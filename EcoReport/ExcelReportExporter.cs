using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

using System.IO;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using DocumentFormat.OpenXml.Wordprocessing;
using Serilog;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

namespace EcoReport
{
	public class ExcelReportExporter
	{
		/// <summary>
		/// 未処理の元のデータ
		/// </summary>
		private List<EcoRenOriginal> _rawData = new List<EcoRenOriginal>();

		/// <summary>
		/// 組あたりのテーブル数
		/// </summary>
		private int _tablesPerGroup = 3;

		/// <summary>
		/// トップスペース
		/// </summary>
		private int _topRows = 0;

		/// <summary>
		/// 表がまたがる行数
		/// </summary>
		private int _tableRowSpan = 12;

		/// <summary>
		/// 表の間隔行
		/// </summary>
		private int _rowGap = 1;

		/// <summary>
		/// 組の間隔行
		/// </summary>
		private int _groupGap = 0;

		/// <summary>
		///  未処理の元のデータ
		/// </summary>
		public List<EcoRenOriginal> RawData
		{
			set { _rawData = value; }
		}

		/// <summary>
		/// 工事名称
		/// </summary>
		public string ConstructionName { get; set; }

		public ExcelReportExporter()
		{
			//ログモジュールの初期化
			GlobalLog.EnsureInitialized();
		}

		private int GetInt(string str)
		{
			var cs = str.ToCharArray();

			for (int i = 0; i < cs.Length; i++)
			{
				var c = cs[i];
				if (char.IsNumber(c) == false) cs[i] = ' ';
			}
			str = new string(cs);
			str.Replace(" ", "");

			var n = Convert.ToInt32(str);
			return n;
		}

		/// <summary>
		///  帳票のエクスポート
		/// </summary>
		public void SaveExcelFile()
		{
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			saveFileDialog.Filter = "Excel ファイル (*.xlsx)|*.xlsx";
			saveFileDialog.DefaultExt = ".xlsx";
			if (saveFileDialog.ShowDialog() == DialogResult.OK)
			{
				var result = ExportReport(saveFileDialog.FileName);
				if (result) MessageBox.Show("エクスポート成功!");
				else MessageBox.Show("エクスポート失敗");
			}
		}

		private XLWorkbook GetTemplateWorkbook()
		{
			try
			{
				var assembly = Assembly.GetExecutingAssembly();
				var resourceName = "EcoReport.template.xlsx"; 

				using (var stream = assembly.GetManifestResourceStream(resourceName))
				{
					if (stream == null)
						throw new FileNotFoundException($"埋め込みリソースが見つかりません: {resourceName}");
					var memoryStream = new MemoryStream();
					stream.CopyTo(memoryStream);
					memoryStream.Seek(0, SeekOrigin.Begin);
					return new XLWorkbook(memoryStream);
				}
			}
			catch (Exception ex)
			{
				GlobalLog.Error(ex, "組み込まれたリソースから Excel テンプレートをロードできなかった。");
				return null;
			}
		}

		/// <summary>
		/// 帳票のエクスポート
		/// </summary>
		private bool ExportReport(string targetPath)
		{
			try
			{
				using (var workbook = GetTemplateWorkbook())
				{
					if (workbook == null)
					{
						GlobalLog.Warning("ファイルテンプレートが見つかりません!");
						return false;
					}

					var templateSheet = workbook.Worksheet("部材別");
					if (templateSheet == null)
					{
						GlobalLog.Warning("ターゲットテンプレート「部材別」は存在しません!");
						return false;
					}

					List<EcoRenInputReport> list = CreateLegalEcoRenInputList();
					List<List<EcoRenInputReport>> grouped = list
						.Select((item, index) => new { item, index })
						.GroupBy(x => x.index / _tablesPerGroup)
						.Select(x => x.Select(o => o.item).ToList())
						.ToList();

					var targetSheet = templateSheet.CopyTo("帳票");
					targetSheet.Clear(XLClearOptions.Contents | XLClearOptions.AllFormats);

					var sourceCell = templateSheet.Range("A1:B1");
					var constructionNameCellHeight = templateSheet.Row(1).Height;
					var tableGapHeight = templateSheet.Row(14).Height;
					var sourceRange = templateSheet.Range("A2:Q13");
					int groupBegin = 1 + _topRows;
					int pageNum = 0;
					foreach (var item in grouped)
					{
						sourceCell.CopyTo(targetSheet.Cell($"A{groupBegin}"));
						//工事名称
						targetSheet.Cell($"B{groupBegin}").Value = ConstructionName;
						targetSheet.Row(groupBegin).Height = constructionNameCellHeight;

						//ページ数
						var cell = targetSheet.Cell($"Q{groupBegin}");
						cell.Style.NumberFormat.Format = "@";
						targetSheet.Cell($"Q{groupBegin}").Value = $"{pageNum + 1}/{grouped.Count}";

						int index = 0;
						groupBegin++;
						int rowBegin = groupBegin;
						int startRow = 0;
						foreach (EcoRenInputReport obj in item)
						{
							startRow = index == 0
							   ? rowBegin + index * _tableRowSpan
							   : rowBegin + index * (_tableRowSpan + _rowGap);

							sourceRange.CopyTo(targetSheet.Cell($"A{startRow}"));

							for (int i = 1; i <= sourceRange.RowCount(); i++)
							{
								int sourceRowNum = sourceRange.FirstRow().RowNumber() + i - 1;
								int targetRowNum = startRow + i - 1;

								var sourceRow = templateSheet.Row(sourceRowNum);
								var targetRow = targetSheet.Row(targetRowNum);
								targetRow.Height = sourceRow.Height;
							}

							FillTableWithData(targetSheet, obj, startRow);
							targetSheet.Row(startRow + _tableRowSpan).Height = tableGapHeight;
							index++;
						}					
						groupBegin = startRow + _tableRowSpan + _groupGap;
						//水平区切り文字の挿入
						targetSheet.PageSetup.AddHorizontalPageBreak(groupBegin-1);
						pageNum++;				
					}
					targetSheet.PageSetup.PrintAreas.Add($"A1:Q{groupBegin - 1}");
					workbook.Worksheets.Delete("部材別");
					workbook.SaveAs(targetPath);
					GlobalLog.Information($"{DateTime.Now.ToString()}:エクスポート成功!");
				}
				return true;
			}
			catch (EcoRenException ex)
			{
				GlobalLog.Error(ex.ToString());
				return false;
			}
			catch (Exception ex)
			{
				GlobalLog.Error($"{DateTime.Now.ToString()}:エクスポート失敗!", ex);
				return false;
			}
		}

		/// <summary>
		/// 合法オブジェクトの作成
		/// </summary>
		/// <returns></returns>
		private List<EcoRenInputReport> CreateLegalEcoRenInputList()
		{
			List<EcoRenInputReport> list = new List<EcoRenInputReport>();
			foreach (var item in _rawData)
			{
				EcoRenInputReport ecoRenInput = new EcoRenInputReport();
				var value = GetInt(item.MainDm);
				if (!ecoRenInput.IsExistsMainDm(value))
				{
					GlobalLog.Warning($"主筋径が正しくありません。主筋径:{value}");
					continue;
				}
				ecoRenInput.mainDm = value;

				value = GetInt(item.MainKind);
				if (!ecoRenInput.IsExistsMainKind(value))
				{
					GlobalLog.Warning($"主筋強度が正しくありません。主筋強度:{value}");
					continue;
				}
				ecoRenInput.mainKind = value;

				value = GetInt(item.StpDm);
				if (!ecoRenInput.IsExistsStpDm(value))
				{
					GlobalLog.Warning($"STP径が正しくありません。STP径:{value}");
					continue;
				}
				ecoRenInput.stpDm = value;

				value = GetInt(item.StpKind);
				if (!ecoRenInput.IsExistsStpKind(value))
				{
					GlobalLog.Warning($"STP強度が正しくありません。STP強度:{value}");
					continue;
				}
				ecoRenInput.stpKind = value;
				ecoRenInput.strListName = item.StrListName;
				ecoRenInput.BodyWidth = GetInt(item.BodyWidth.ToString());
				ecoRenInput.BodyHeight = GetInt(item.BodyHeight.ToString());
				ecoRenInput.Fc = GetInt(item.Fc.ToString());
				ecoRenInput.Kaburi = GetInt(item.Kaburi.ToString());
				ecoRenInput.LeftHiCount = GetInt(item.LeftHiCount.ToString());
				ecoRenInput.LeftLowCount = GetInt(item.LeftLowCount.ToString());
				ecoRenInput.RightHiCount = GetInt(item.RightHiCount.ToString());
				ecoRenInput.RightLowCount = GetInt(item.RightLowCount.ToString());
				ecoRenInput.stpCount = GetInt(item.StpCount);
				ecoRenInput.stpPitch = GetInt(item.StpPitch);
				ecoRenInput.Uchinori = GetInt(item.Uchinori.ToString());
				list.Add(ecoRenInput);
			}
			return list;
		}

		/// <summary>
		/// データで表を埋める
		/// </summary>
		/// <param name="sheet">表の場所sheet</param>
		/// <param name="data">データソース</param>
		/// <param name="startRow">開始行</param>
		private void FillTableWithData(IXLWorksheet sheet, EcoRenInputReport ecoRenInput, int startRow)
		{
			//階：梁記号
			sheet.Cell($"B{startRow}").Value = ecoRenInput.strFloor;
			sheet.Cell($"C{startRow}").Value = ecoRenInput.strListName;

			int lineSpacing = 2;
			#region  条件
			//b(㎜)
			sheet.Cell($"B{startRow + lineSpacing}").Value = ecoRenInput.BodyWidth;
			//D(㎜)
			sheet.Cell($"C{startRow + lineSpacing}").Value = ecoRenInput.BodyHeight;
			//d(㎜)
			sheet.Cell($"D{startRow + lineSpacing}").Value = ecoRenInput.span_d;
			//C(㎜)
			sheet.Cell($"E{startRow + lineSpacing}").Value = ecoRenInput.rangeC_C;
			//Fc(N/㎟)
			sheet.Cell($"F{startRow + lineSpacing}").Value = ecoRenInput.Fc;
			//ｌ’(㎜)
			sheet.Cell($"G{startRow + lineSpacing}").Value = ecoRenInput.Uchinori;
			//Qo
			sheet.Cell($"H{startRow + lineSpacing}").Value = ecoRenInput.Qo;
			//β
			sheet.Cell($"I{startRow + lineSpacing}").Value = Math.Min(ecoRenInput.shukyoku_QSU_B, ecoRenInput.shukyoku_Qm_B);
			//上端筋
			sheet.Cell($"K{startRow + lineSpacing}").Value = $"{ecoRenInput.LeftHiCount}ー{ecoRenInput.mainDm}";
			//下端筋
			sheet.Cell($"L{startRow + lineSpacing}").Value = $"{ecoRenInput.LeftLowCount}ー{ecoRenInput.mainDm}";
			//上端筋
			sheet.Cell($"M{startRow + lineSpacing}").Value = $"{ecoRenInput.RightHiCount}ー{ecoRenInput.mainDm}";
			//下端筋
			sheet.Cell($"N{startRow + lineSpacing}").Value = $"{ecoRenInput.RightLowCount}ー{ecoRenInput.mainDm}";
			//鋼種
			sheet.Cell($"O{startRow + lineSpacing}").Value = ecoRenInput.mainKind;
			//せん断補強筋
			sheet.Cell($"P{startRow + lineSpacing}").Value = $"{ecoRenInput.stpCount}ー{ecoRenInput.stpDm}";
			sheet.Cell($"Q{startRow + lineSpacing}").Value = ecoRenInput.stpPitch;
			#endregion

			lineSpacing = 4;
			#region 無孔梁
			//Pt
			sheet.Cell($"B{startRow + lineSpacing}").Value = ecoRenInput.span_max_Pt;
			//Pw
			sheet.Cell($"C{startRow + lineSpacing}").Value = ecoRenInput.shukyoku_Pw;
			//ΣMy
			sheet.Cell($"D{startRow + lineSpacing}").Value = ecoRenInput.design_sendan_HiMy;
			//ΣMy/l'
			sheet.Cell($"E{startRow + lineSpacing}").Value = ecoRenInput.design_sendan_HiMy / ecoRenInput.Uchinori;
			//maxMy
			sheet.Cell($"F{startRow + lineSpacing}").Value = ecoRenInput.design_sendan_max_My;
			//M/Qd
			sheet.Cell($"G{startRow + lineSpacing}").Value = ecoRenInput.design_sendan_MQd2;
			//uτD
			sheet.Cell($"H{startRow + lineSpacing}").Value = ecoRenInput.shukyoku_design_sendan_tmin;
			//τmu
			sheet.Cell($"I{startRow + lineSpacing}").Value = ecoRenInput.design_sendan_tmu;
			//uτDmin
			sheet.Cell($"J{startRow + lineSpacing}").Value = Math.Min(ecoRenInput.design_sendan_tmu, ecoRenInput.shukyoku_design_sendan_tmin);
			//uτmin/β
			sheet.Cell($"K{startRow + lineSpacing}").Value = Math.Min(ecoRenInput.design_sendan_tmu, ecoRenInput.shukyoku_design_sendan_tmin) / Math.Min(ecoRenInput.shukyoku_QSU_B, ecoRenInput.shukyoku_Qm_B);
			#endregion

			lineSpacing = 6;
			#region 開孔補強
			#region Φ100
			try
			{
				//S-エコレン
				sheet.Cell($"B{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"B{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"B{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"B{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"B{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"B{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "B", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ125
			try
			{
				//S-エコレン
				sheet.Cell($"D{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"D{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"D{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"D{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"D{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"D{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "D", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ150
			try
			{
				//S-エコレン
				sheet.Cell($"F{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"F{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"F{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"F{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"F{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"F{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "F", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ175
			try
			{
				//S-エコレン
				sheet.Cell($"H{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"H{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"H{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"H{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"H{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"H{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "H", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ200
			try
			{
				//S-エコレン
				sheet.Cell($"J{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"J{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"J{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"J{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"J{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"J{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "J", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ250
			try
			{
				//S-エコレン
				sheet.Cell($"L{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"L{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"L{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"L{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"L{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"L{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "L", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ300
			try
			{
				//S-エコレン
				sheet.Cell($"N{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"N{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"N{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"N{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"N{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"N{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "N", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion

			#region Φ350
			try
			{
				//S-エコレン
				sheet.Cell($"P{startRow + lineSpacing}").Value = ecoRenInput.shukyoku;
				//Pe
				sheet.Cell($"P{startRow + lineSpacing + 1}").Value = ecoRenInput.shukyoku_QSU_ePw;
				//Ps
				sheet.Cell($"P{startRow + lineSpacing + 2}").Value = ecoRenInput.shukyoku_S_hokyou_sPw;
				//aPe(mm2)
				sheet.Cell($"P{startRow + lineSpacing + 3}").Value = ecoRenInput.shukyoku_QSU_S;
				//uτH
				sheet.Cell($"P{startRow + lineSpacing + 4}").Value = ecoRenInput.shukyoku_QSU_utH;
				//uτH/uτD
				sheet.Cell($"P{startRow + lineSpacing + 5}").Value = ecoRenInput.shukyoku_QSU_utHutD;
			}
			catch (EcoRenException ex)
			{
				SetTableColumnEmpty(sheet, "P", startRow + lineSpacing);
				GlobalLog.Warning($"{ex.Code}:{ex.strMessage}");
			}
			#endregion
			#endregion
		}

		/// <summary>
		/// カラム内容をNULLに設定
		/// </summary>
		/// <param name="sheet">sheet</param>
		/// <param name="columnChar">列文字</param>
		/// <param name="startRow">行番号</param>
		private void SetTableColumnEmpty(IXLWorksheet sheet, string columnChar, int startRow)
		{
			//S-エコレン
			sheet.Cell($"{columnChar}{startRow}").Value = string.Empty;
			//Pe
			sheet.Cell($"{columnChar}{startRow + 1}").Value = string.Empty;
			//Ps								  
			sheet.Cell($"{columnChar}{startRow + 2}").Value = string.Empty;
			//aPe(mm2)							  
			sheet.Cell($"{columnChar}{startRow + 3}").Value = string.Empty;
			//uτH								   
			sheet.Cell($"{columnChar}{startRow + 4}").Value = string.Empty;
			//uτH/uτD							  
			sheet.Cell($"{columnChar}{startRow + 5}").Value = string.Empty;
		}
	}
}
