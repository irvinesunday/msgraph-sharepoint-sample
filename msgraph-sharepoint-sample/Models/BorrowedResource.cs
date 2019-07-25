using Newtonsoft.Json;
using System;

namespace msgraph_sharepoint_sample.Models
{
    class BorrowedResource
    {
        [JsonIgnore]
        public string SharePointItemId { get; set; }

        [JsonProperty(PropertyName = "BorrowedResourceId")]

        public Guid BorrowedResourceId { get; set; }
        
        [JsonProperty(PropertyName = "Resource_x0020_IDLookupId")]
        public string ResourceId { get; set; }

        public string ResourceStatusId { get; set; }

    }
}
