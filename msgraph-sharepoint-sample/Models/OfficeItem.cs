using Newtonsoft.Json;
using System;

namespace msgraph_sharepoint_sample.Models
{
    class OfficeItem
    {
        [JsonIgnore]
        public string SharePointItemId { get; set; }

        [JsonProperty(PropertyName = "OfficeItemID")]

        public Guid ItemId { get; set; }

        public string Title { get; set; }

        [JsonProperty(PropertyName = "Resource_x0020_IDLookupId")]
        public string ResourceId { get; set; }

        [JsonProperty(PropertyName = "SerialNo")]

        public string SerialNumber { get; set; }

        [JsonProperty(PropertyName = "Description")]
        public string ItemDescription { get; set; }

    }
}
