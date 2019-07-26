using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample.Models
{
    public class Member
    {
        [JsonIgnore]
        public string SharePointItemId { get; set; }
        
        [JsonProperty(PropertyName = "MemberID")]
        public Guid MemberId { get; set; }
        [JsonProperty(PropertyName = "Title")]
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        [JsonProperty(PropertyName = "isAdmin")]
        public string IsAdmin { get; set; }
        
    }
}
