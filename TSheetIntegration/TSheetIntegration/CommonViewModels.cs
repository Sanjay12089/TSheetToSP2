using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSheetIntegration
{
    public class CommonViewModels
    {

    }

    public class AllTimeSheetData
    {
        public long id { get; set; }
        public long user_id { get; set; }
        public long jobcode_id { get; set; }
        public DateTime? start { get; set; }
        public DateTime? end { get; set; }
        public string duration { get; set; }
        public DateTime? date { get; set; }
        public string tz { get; set; }
        public string tz_str { get; set; }
        public string type { get; set; }
        public string location { get; set; }
        public string on_the_clock { get; set; }
        public string locked { get; set; }
        public string notes { get; set; }
        public CustomFields customfields { get; set; }
        public string last_modified { get; set; }
        //public byte[] attached_files { get; set; }
        public string created_by_user_id { get; set; }
    }

    public class SupplementalData
    {
        public long id { get; set; }
        public long parent_id { get; set; }
        public bool assigned_to_all { get; set; }
        public bool billable { get; set; }
        public bool active { get; set; }
        public string type { get; set; }
        public bool has_children { get; set; }
        public string billable_rate { get; set; }
        public string short_code { get; set; }
        public string name { get; set; }
        public string last_modified { get; set; }
        public DateTime created { get; set; }
        //public byte[] required_customfields { get; set; }
        //public byte[] locations { get; set; }
        public long geofence_config_id { get; set; }
        public long project_id { get; set; }
    }

    public class MilestoneData
    {
        public long id { get; set; }
        public string name { get; set; }
    }

    public class CustomFields
    {
        [Column("32318")]
        public string FirstColumn { get; set; }
        [Column("78682")]
        public string SecondColumn { get; set; }
        [Column("32316")]
        public string ThirdColumn { get; set; }
        [Column("78928")]
        public string FourthColumn { get; set; }
        [Column("75755")]
        public string FifthColumn { get; set; }
    }
}