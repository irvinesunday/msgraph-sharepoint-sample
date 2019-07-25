using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample.Models
{
    public class Member
    {
        public string SharePointItemId { get; set; }
        public Guid MemberId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public bool IsAdmin { get; set; }
        public DateTime DateCreated { get; set; }
    }
}
