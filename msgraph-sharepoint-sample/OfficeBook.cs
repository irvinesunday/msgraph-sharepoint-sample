using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample
{
    public class OfficeBook
    {
        [JsonIgnore]
        public string SharePointItemId { get; set; }

        public Guid BookId { get; set; }

        public string Title { get; set; }

        [JsonProperty(PropertyName = "Resource_x0020_IDLookupId")]
        public string ResourceId { get; set; }

        public string ISBN { get; set; }

        [JsonProperty(PropertyName = "BookTitle")]
        public string BookDescription { get; set; }

        [JsonProperty(PropertyName = "Author0")]
        public string Author { get; set; }
    }
}
