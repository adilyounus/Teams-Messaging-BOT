using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.MSGraphModels
{
    public class CreateChatModel
    {
        public string chatType { get; set; }
        public List<ChatMemberModel> members { get; set; }
    }

    public class ChatMemberModel
    {
        public ChatMemberModel()
        {
        }

        public ChatMemberModel(string user)
        {
            userodatabind = user;
        }

        [JsonProperty("@odata.type")]
        public string odatatype { get; set; } = "#microsoft.graph.aadUserConversationMember";
        public string[] roles { get; set; } = new string[] { "owner" };

        [JsonProperty("user@odata.bind")]
        public string userodatabind { get; set; }
    }

    public class ChatMessageModel
    {
        public ChatMessageBodyModel body { get; set; }
        public List<ChatMessageMentionModel> mentions { get; set; }
        public List<ChatMessageHostedContentModel> hostedContents { get; set; }
    }

    public class ChatMessageBodyModel
    {
        public string contentType { get; set; } = "html";
        public string content { get; set; }
    }

    public class ChatMessageMentionModel
    {
        public int id { get; set; }
        public string mentionText { get; set; }
        public ChatMessageMentionedModel mentioned { get; set; }
    }

    public class ChatMessageMentionedModel
    {
        public ChatMessageUserModel user { get; set; }
    }

    public class ChatMessageMentionTagModel
    {
        public int id { get; set; }
        public string mentionText { get; set; }
        public ChatMessageMentionedTagModel mentioned { get; set; }
    }
    public class ChatMessageMentionedTagModel
    {
        public ChatMessageTagModel tag { get; set; }
    }

    public class ChatMessageTagModel
    {
        public string id { get; set; }
        public string displayName { get; set; }
    }


    public class ChatMessageUserModel
    {
        public string displayName { get; set; }
        public string id { get; set; }
        public string userIdentityType { get; set; }
    }

    public class ChatMessageHostedContentModel
    {
        [JsonProperty("@microsoft.graph.temporaryId")]
        public string MicrosoftGraphTemporaryId { get; set; }
        public string contentBytes { get; set; }
        public string contentType { get; set; }
    }

    public class ChatMessageAttachmentModel
    {
        public string id { get; set; }
        public string contentType { get; set; } = "application/vnd.microsoft.card.thumbnail";
        public object contentUrl { get; set; } = null;
        public string content { get; set; }
        public object name { get; set; } = null;
        public object thumbnailUrl { get; set; } = null;
    }

    public class ChatMessageAttachmentContentMoel
    {
        public string title { get; set; }
        public string subtitle { get; set; }
        public string text { get; set; }
    }

}
