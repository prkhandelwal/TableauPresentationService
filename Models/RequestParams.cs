/*
 * Created By Pratik Khandelwal
 */
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Web;

namespace PresentationService.Models
{
    public class RequestParams
    {
        public List<string> dashboardList { get; private set; }
        public string dashboard
        {
            get
            {
                return "Dashboards";
            }
            set
            {
                dashboardList = value.Split(",").ToList();
            }
        }
        public string globalFilters { get; set; }

    }
}