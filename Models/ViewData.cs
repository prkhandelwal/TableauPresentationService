/*
 * Created By Pratik Khandelwal
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PresentationService.Models
{
    public class ViewData
    {
        public ViewData(string name, string viewId, 
            string chartType, string primaryAxis,
            List<string> serieList)
        {
            this.Name= name;
            this.ViewId = viewId;
            this.ChartType = chartType;
            this.PrimaryAxis = primaryAxis;
            this.SerieList = serieList;
        }
        public string Name { get; set; }
        public string ViewId { get; set; }
        public Stream DataStream { get; set; }
        public string ChartType { get; set; }
        public string PrimaryAxis { get; set; }
        public List<string> SerieList { get; set; }
    }
}
