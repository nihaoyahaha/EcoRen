using EcoReport;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EcoRen_MainForm
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}
		private List<EcoRenOriginal> InitDataSource()
		{
			return new List<EcoRenOriginal>()
			{
				 new EcoRenOriginal()
			   {
				StrListName = "FG1",
				BodyWidth = 800,
				BodyHeight = 3500,
				Fc = 36,
				Kaburi = 140,
				MainDm = "D35",
				LeftHiCount = 10,
				LeftLowCount = 10,
				RightHiCount = 10,
				RightLowCount = 10,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@150",
				StpKind = "SD295",
				Uchinori = 5050
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
				 new EcoRenOriginal()
				 {
				StrListName = "FG2",
				BodyWidth = 700,
				BodyHeight = 2250,
				Fc = 36,
				Kaburi = 100,
				MainDm = "D35",
				LeftHiCount = 7,
				LeftLowCount = 7,
				RightHiCount = 7,
				RightLowCount = 7,
				MainKind = "SD390",
				StpCount = "2-",
				StpDm = "D16",
				StpPitch = "@200",
				StpKind = "SD295",
				Uchinori = 5200
			   },
			};
		}

		private void button1_Click(object sender, EventArgs e)
		{
			ExcelReportExporter excelReport = new ExcelReportExporter();
			excelReport.RawData = InitDataSource();
			excelReport.ConstructionName = "工事名称";
			excelReport.SaveExcelFile();
		}
	}
}
