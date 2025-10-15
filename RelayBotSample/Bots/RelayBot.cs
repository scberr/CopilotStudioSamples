// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Bot.Builder;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Connector.DirectLine;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DirectLineActivity = Microsoft.Bot.Connector.DirectLine.Activity;
using DirectLineActivityTypes = Microsoft.Bot.Connector.DirectLine.ActivityTypes;
using IConversationUpdateActivity = Microsoft.Bot.Schema.IConversationUpdateActivity;
using IMessageActivity = Microsoft.Bot.Schema.IMessageActivity;

namespace Microsoft.PowerVirtualAgents.Samples.RelayBotSample.Bots
{
    /// <summary>
    /// This IBot implementation shows how to connect
    /// an external Azure Bot Service channel bot (external bot)
    /// to your Power Virtual Agent bot
    /// </summary>
    public class RelayBot : ActivityHandler
    {
        private const int WaitForBotResponseMaxMilSec = 10 * 1000;
        private const int PollForBotResponseIntervalMilSec = 1000;
        private static ConversationManager s_conversationManager = ConversationManager.Instance;
        private ResponseConverter _responseConverter;
        private IBotService _botService;
        private ILogger<RelayBot> _logger;

        public RelayBot(IBotService botService, ConversationManager conversationManager, ILogger<RelayBot> logger)
        {
            _botService = botService;
            _responseConverter = new ResponseConverter();
            _logger = logger;
        }

        // Invoked when a conversation update activity is received from the external Azure Bot Service channel
        // Start a Power Virtual Agents bot conversation and store the mapping
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            await s_conversationManager.GetOrCreateBotConversationAsync(turnContext.Activity.Conversation.Id, _botService);
        }

        // Invoked when a message activity is received from the user
        // Send the user message to Power Virtual Agent bot and get response
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var currentConversation = await s_conversationManager.GetOrCreateBotConversationAsync(turnContext.Activity.Conversation.Id, _botService);

            _logger.LogInformation($"Processing message activity from user: {turnContext.Activity.From.Id}, conversationId: {turnContext.Activity.Conversation.Id}, text: {turnContext.Activity.Text}");

            using (DirectLineClient client = new DirectLineClient(currentConversation.Token))
            {
                // Send user message using directlineClient
                var response = await client.Conversations.PostActivityAsync(currentConversation.ConversationtId, new DirectLineActivity()
                {
                    Type = DirectLineActivityTypes.Message,
                    From = new ChannelAccount { Id = turnContext.Activity.From.Id, Name = turnContext.Activity.From.Name },
                    Text = turnContext.Activity.Text,
                    TextFormat = turnContext.Activity.TextFormat,
                    Locale = turnContext.Activity.Locale,
                });

                await RespondCopilotStudioAgentReplyAsync(client, currentConversation, turnContext);
            }

            // Update LastConversationUpdateTime for session management
            currentConversation.LastConversationUpdateTime = DateTime.Now;
        }

        private async Task RespondCopilotStudioAgentReplyAsync(DirectLineClient client, RelayConversation currentConversation, ITurnContext<IMessageActivity> turnContext)
        {

            var retryMax = WaitForBotResponseMaxMilSec / PollForBotResponseIntervalMilSec;

            using (_logger.BeginScope(new Dictionary<string, object>
            {
                ["ConversationId"] = turnContext.Activity.Conversation.Id,
                ["Method"] = "Copilot Studio Response"
            }))
            {

                _logger.LogInformation($"Waiting for response");

                for (int retry = 0; retry < retryMax; retry++)
                {

                    _logger.LogInformation($"Retry #{retry + 1} initiated");

                    // Get bot response using directlineClient,
                    // response contains whole conversation history including user & bot's message
                    ActivitySet response = await client.Conversations.GetActivitiesAsync(currentConversation.ConversationtId, currentConversation.WaterMark);

                    // Filter bot's reply message from response
                    List<DirectLineActivity> botResponses = response?.Activities?.Where(x =>
                          x.Type == DirectLineActivityTypes.Message &&
                            string.Equals(x.From.Name, _botService.GetBotName(), StringComparison.Ordinal)).ToList();

                    if (botResponses?.Count() > 0)
                    {
                        if (int.Parse(response?.Watermark ?? "0") <= int.Parse(currentConversation.WaterMark ?? "0"))
                        {
                            // means user sends new message, should break previous response poll
                            return;
                        }

                        _logger.LogInformation($"Response Received!");

                        currentConversation.WaterMark = response.Watermark;
                        var responses = _responseConverter.ConvertToBotSchemaActivities(botResponses).ToArray();
                        _logger.LogInformation($"Sending {responses.Length} responses back to client");
                        await turnContext.SendActivitiesAsync(responses);
                    }

                    Thread.Sleep(PollForBotResponseIntervalMilSec);

                    _logger.LogInformation($"Retry #{retry+1} failed");
                }

                _logger.LogInformation($"No responses received, polling for new messages ended.");

            }
        }
    }
}
