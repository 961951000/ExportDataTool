using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExportDataTool.Models
{
    public class ExportModel
    {
        public string DataSource { get; set; }

        [Required]
        public string ConnectionString { get; set; }

        [Required]
        public string SQL { get; set; }
    }
}
