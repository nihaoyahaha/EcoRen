using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EcoReport
{
	public class ProjectConfiguration
	{
		public string TemplatePath { get; set; }
		public string LogMinimumLevel { get; set; }
		public string LogPath { get; set; }
		public int LogFileSizeLimitBytes { get; set; }
		public bool LogRollOnFileSizeLimit { get; set; }
		public RollingInterval LogRollingInterval { get; set; }
		public int LogRetainedFileCountLimit { get; set; }
	}
}
