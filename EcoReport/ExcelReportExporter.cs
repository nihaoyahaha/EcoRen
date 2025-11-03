using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using Serilog;
using DocumentFormat.OpenXml.Spreadsheet;
using EcoRecDatas;
using System.Drawing;

namespace EcoReport
{
	public class ExcelReportExporter
	{
		/// <summary>
		/// テンプレートパス
		/// </summary>
		private string _templatePath;

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
			_templatePath = GlobalLog.DeserializeProjectConfiguration().TemplatePath;
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

		/// <summary>
		/// 帳票のエクスポート
		/// </summary>
		private bool ExportReport(string targetPath)
		{
			try
			{
				if (!File.Exists(_templatePath))
				{
					GlobalLog.Warning("ファイルテンプレートが見つかりません!");
					return false;
				}
				List<EcoRenInputReport> list = CreateLegalEcoRenInputList();
				List<List<EcoRenInputReport>> grouped = list
					.Select((item, index) => new { item, index })
					.GroupBy(x => x.index / _tablesPerGroup)
					.Select(x => x.Select(o => o.item).ToList())
					.ToList();

				File.Copy(_templatePath, targetPath, true);

				using (var workbook = new XLWorkbook(targetPath))
				{
					var templateSheet = workbook.Worksheet("部材別");
					if (templateSheet == null)
					{
						GlobalLog.Warning("ターゲットテンプレート「部材別」は存在しません!");
						return false;
					}
					templateSheet = workbook.Worksheet("計算例");
					if (templateSheet == null)
					{
						GlobalLog.Warning("ターゲットテンプレート「計算例」は存在しません!");
						return false;
					}

					//計算例Sheetの作成
					if (list.Count > 0)
					{
						CreateCalculationExampleSheet(workbook, list[0]);
					}

					//部材別Sheetの作成
					CreateBuzaiSheet(workbook, grouped);

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
		/// 部材別Sheetの作成
		/// </summary>
		/// <param name="workbook">ワークブック</param>
		/// <param name="grouped">グループ化されたデータ</param>
		/// <returns>部材別Sheet</returns>
		private void CreateBuzaiSheet(XLWorkbook workbook, List<List<EcoRenInputReport>> grouped)
		{
			var templateSheet = workbook.Worksheet("部材別");
			var targetSheet = templateSheet.CopyTo("sheet1");
			targetSheet.Clear(XLClearOptions.Contents | XLClearOptions.AllFormats);

			var pageStyle = templateSheet.Cell("Q1").Style;
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
				targetSheet.Cell($"Q{groupBegin}").Style = pageStyle;

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

					FillBuzaiSheetWithData(targetSheet, obj, startRow);
					targetSheet.Row(startRow + _tableRowSpan).Height = tableGapHeight;
					index++;
				}
				groupBegin = startRow + _tableRowSpan + _groupGap;
				//水平区切り文字の挿入
				targetSheet.PageSetup.AddHorizontalPageBreak(groupBegin - 1);

				pageNum++;
			}
			targetSheet.PageSetup.PrintAreas.Add($"A1:Q{groupBegin - 1}");
			workbook.Worksheets.Delete("部材別");
			targetSheet.Name = "部材別";
			workbook.Save();			
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
		private void FillBuzaiSheetWithData(IXLWorksheet sheet, EcoRenInputReport ecoRenInput, int startRow)
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

		/// <summary>
		/// 構成ファイルからテンプレートパスを取得するには
		/// </summary>
		/// <returns></returns>
		private string GetTemplateFilePath()
		{
			return GlobalLog.DeserializeProjectConfiguration().TemplatePath;
		}

		/// <summary>
		///  計算例Sheetの作成
		/// </summary>
		/// <param name="workbook">ワークブック</param>
		/// <param name="ecoRenReport">データ</param>
		private void CreateCalculationExampleSheet(XLWorkbook workbook, EcoRenInputReport ecoRenRpt)
		{
			var targetSheet = workbook.Worksheet("計算例");

			//工事名称
			targetSheet.Cell("B1").Value = ConstructionName;
			//階
			targetSheet.Cell("D2").Value = ecoRenRpt.strFloor;
			//梁記号
			targetSheet.Cell("F2").Value = ecoRenRpt.strListName;
			//開孔径φ
			targetSheet.Cell("C3").Value = ecoRenRpt.PAI;

			#region 表
			//b(㎜)
			targetSheet.Cell("A5").Value = ecoRenRpt.BodyWidth;
			//D(㎜)
			targetSheet.Cell("B5").Value = ecoRenRpt.BodyHeight;
			//d(㎜)
			targetSheet.Cell("C5").Value = ecoRenRpt.span_d;
			//dc(㎜)
			targetSheet.Cell("D5").Value = ecoRenRpt.span_dc;
			//C(㎜)
			targetSheet.Cell("E5").Value = ecoRenRpt.rangeC_C;
			//Fc(N/㎟)
			targetSheet.Cell("F5").Value = ecoRenRpt.Fc;
			//ｌ’(㎜)
			targetSheet.Cell("G5").Value = ecoRenRpt.Uchinori;
			//Q0
			targetSheet.Cell("H5").Value = ecoRenRpt.Qo;
			//β
			targetSheet.Cell("I5").Value = Math.Min(ecoRenRpt.shukyoku_QSU_B, ecoRenRpt.shukyoku_Qm_B);
			//上端筋
			targetSheet.Cell("K5").Value = $"{ecoRenRpt.LeftHiCount}ー{ecoRenRpt.mainDm}";
			//下端筋
			targetSheet.Cell("L5").Value = $"{ecoRenRpt.LeftLowCount}ー{ecoRenRpt.mainDm}";
			//上端筋
			targetSheet.Cell("M5").Value = $"{ecoRenRpt.RightHiCount}ー{ecoRenRpt.mainDm}";
			//下端筋
			targetSheet.Cell("N5").Value = $"{ecoRenRpt.RightLowCount}ー{ecoRenRpt.mainDm}";
			//鋼種
			targetSheet.Cell("O5").Value = ecoRenRpt.mainKind;
			//せん断補強筋
			targetSheet.Cell("P5").Value = $"{ecoRenRpt.stpCount}ー{ecoRenRpt.stpDm}";
			targetSheet.Cell("Q5").Value = ecoRenRpt.stpPitch;
			//Pt
			targetSheet.Cell("A7").Value = ecoRenRpt.span_max_Pt;
			//Pw
			targetSheet.Cell("B7").Value = ecoRenRpt.shukyoku_Pw;
			//ΣMy
			targetSheet.Cell("C7").Value = ecoRenRpt.design_sendan_max_Qmu;
			//M/Qd
			targetSheet.Cell("D7").Value = ecoRenRpt.design_sendan_MQd2;
			//τmu
			targetSheet.Cell("E7").Value = ecoRenRpt.design_sendan_tmu;
			//uτD
			targetSheet.Cell("F7").Value = ecoRenRpt.shukyoku_design_sendan_tmin;
			//uτD min
			targetSheet.Cell("G7").Value = Math.Min(ecoRenRpt.design_sendan_tmu, ecoRenRpt.shukyoku_design_sendan_tmin);
			//uτD min/β
			targetSheet.Cell("H7").Value = Math.Min(ecoRenRpt.design_sendan_tmu, ecoRenRpt.shukyoku_design_sendan_tmin) / Math.Min(ecoRenRpt.shukyoku_QSU_B, ecoRenRpt.shukyoku_Qm_B);
			#endregion

			#region テキスト
			//Pt=(at/(b･d))/100=(span_right_Count･span_right_S/(BodyWidth･span_d))･100=span_max_Pt  (%) 
			var cell = targetSheet.Cell("B9");
			var richText = cell.GetRichText();
			richText.ClearText()
				.AddText("Pt=(at/(b･d))/100=(")
				.AddText($"{ecoRenRpt.span_right_Count}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.span_right_S}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/(")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.span_d}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("))･100=")
				.AddText($"{ecoRenRpt.span_max_Pt}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("  (%) ");

			//Pw=aw/(b･x)=(stpCount･sendan_hokyo_S)/(BodyWidth･stpPitch)=shukyoku_Pw
			cell = targetSheet.Cell("B10");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("Pw=aw/(b･x)=(")
				.AddText($"{ecoRenRpt.stpCount}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.sendan_hokyo_S}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")/(")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.stpPitch}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")=")
				.AddText($"{ecoRenRpt.shukyoku_Pw}").SetFontColor(XLColor.Red).SetBold(true);

			//ΣMy=上My+下My=上(0.9･at･1.1･σy･d)+下(0.9･at･1.1･σy･d)=0.9･(span_left_at/1000)･1.1･mainKind･(span_d/1000)+0.9･(span_right_at/1000)･1.1･mainKind･(span_d/1000)=design_sendan_LowMy+design_sendan_My=design_sendan_HiMy(kN・m)
			cell = targetSheet.Cell("B11");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("ΣMy=上My+下My=上(0.9･at･1.1･σy･d)+下(0.9･at･1.1･σy･d)=0.9･(")
				.AddText($"{ecoRenRpt.span_left_at}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)･1.1･")
				.AddText($"{ecoRenRpt.mainKind}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･(")
				.AddText($"{ecoRenRpt.span_d}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)+0.9･(")
				.AddText($"{ecoRenRpt.span_right_at}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)･1.1･")
				.AddText($"{ecoRenRpt.mainKind}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･(")
				.AddText($"{ecoRenRpt.span_d}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)=")
				.AddText($"{ecoRenRpt.design_sendan_LowMy}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+")
				.AddText($"{ecoRenRpt.design_sendan_My}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("=")
				.AddText($"{ecoRenRpt.design_sendan_HiMy}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("(kN・m)");

			//My=MAX(design_sendan_LowMy,design_sendan_My)=design_sendan_max_My(kN･m)
			cell = targetSheet.Cell("B12");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("My=MAX(")
				.AddText($"{ecoRenRpt.design_sendan_LowMy}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(",")
				.AddText($"{ecoRenRpt.design_sendan_My}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText($")=")
				.AddText($"{ecoRenRpt.design_sendan_max_My}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("(kN･m)");

			//M/Qd=maxMy/(Qm･d)=design_sendan_max_My/(design_sendan_max_Qmu･(span_d/1000))=design_sendan_MQd2　(1≦M/Q･ｄ≦3)
			cell = targetSheet.Cell("B13");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("M/Qd=maxMy/(Qm･d)=")
				.AddText($"{ecoRenRpt.design_sendan_max_My}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/(")
				.AddText($"{ecoRenRpt.design_sendan_max_Qmu}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･(")
				.AddText($"{ecoRenRpt.span_d}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000))=")
				.AddText($"{ecoRenRpt.design_sendan_MQd2}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("　(1≦M/Q･ｄ≦3)");

			//C=D/2-dc= BodyHeight-2･ span_dc=rangeC_C (㎜)
			cell = targetSheet.Cell("B14");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("C=D/2-dc= ")
				.AddText($"{ecoRenRpt.BodyHeight}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("-2･ ")
				.AddText($"{ecoRenRpt.span_dc}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("=")
				.AddText($"{ecoRenRpt.rangeC_C}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" (㎜)");

			//sPw=saw/(b･C区間)=rangeC_Cst･stpCount･sendan_hokyo_S/(BodyWidth･rangeC_C)=shukyoku_S_hokyou_sPw
			cell = targetSheet.Cell("B15");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("sPw=saw/(b･C区間)=")
				.AddText($"{ecoRenRpt.rangeC_Cst}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.stpCount}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.sendan_hokyo_S}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/(")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.rangeC_C}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")=")
				.AddText($"{ecoRenRpt.shukyoku_S_hokyou_sPw}").SetFontColor(XLColor.Red).SetBold(true);

			#region 設計せん断力
			//両端曲げ降伏
			//Qm=Qo＋n･ΣMy/l'=Qo+[design_sendan_a･design_sendan_HiMy/(Uchinori/1000)]=design_sendan_max_Qmu (kN)
			cell = targetSheet.Cell("B18");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("Qm=Qo＋n･ΣMy/l'=")
				.AddText($"{ecoRenRpt.Qo}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+[")
				.AddText($"{ecoRenRpt.design_sendan_a}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.design_sendan_HiMy}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/(")
				.AddText($"{ecoRenRpt.Uchinori}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)]=")
				.AddText($"{ecoRenRpt.design_sendan_max_Qmu}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" (kN)");

			//両端曲げ降伏
			//τｍu=Qm/(b･j)=design_sendan_max_Qmu/(BodyWidth･span_j/1000)=design_sendan_tmu (N/㎟)
			cell = targetSheet.Cell("B19");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("τｍu=Qm/(b･j)=")
				.AddText($"{ecoRenRpt.design_sendan_max_Qmu}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/(")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.span_j}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)=")
				.AddText($"{ecoRenRpt.design_sendan_tmu}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" (N/㎟)");

			//荒川min
			//QD={0.053･Pt0.23･(Fc+18)/[(M/Qd)+0.12]+0.85･(Pw･sσy)0.5}･ｂ･j={[0.053･(span_max_Pt0.23)･(Fc+18))/(design_sendan_MQd2+0.12)]+0.85･(shukyoku_Pw･stpKind)0.5}･BodyWidth･span_j/1000=shukyoku_design_sendan_Qmin (kN)
			cell = targetSheet.Cell("B21");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("Q")
				.AddText("D").VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("={0.053･Pt")
				.AddText("0.23").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText("･(Fc+18)/[(M/Qd)+0.12]+0.85･(Pw･sσy)")
				.AddText("0.5").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText("}･ｂ･j={[0.053･(")
				.AddText($"{ecoRenRpt.span_max_Pt}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("0.23").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText(")･(")
				.AddText($"{ecoRenRpt.Fc}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+18))/(")
				.AddText($"{ecoRenRpt.design_sendan_MQd2}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+0.12)]+0.85･(")
				.AddText($"{ecoRenRpt.shukyoku_Pw}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.stpKind}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")")
				.AddText("0.5").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText("}･")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.span_j}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000=")
				.AddText($"{ecoRenRpt.shukyoku_design_sendan_Qmin}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" (kN)");

			//荒川min
			//uτD=QD/(ｂ･j)=shukyoku_design_sendan_Qmin/(BodyWidth･span_j/1000)=shukyoku_design_sendan_tmin (N/㎟)
			cell = targetSheet.Cell("B22");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("uτ")
				.AddText("D").VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("=Q")
				.AddText("D").VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("/(ｂ･j)=")
				.AddText($"{ecoRenRpt.shukyoku_design_sendan_Qmin}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/(")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.span_j}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/1000)=")
				.AddText($"{ecoRenRpt.shukyoku_design_sendan_tmin}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" (N/㎟)");

			//設計せん断力
			//uτD=min(τmu、uτD)=mim (design_sendan_tmu、shukyoku_design_sendan_tmin)=min(design_sendan_tmu,shukyoku_design_sendan_tmin)
			cell = targetSheet.Cell("B24");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("uτ")
				.AddText("D").VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("=min(τmu、uτ")
				.AddText("D").VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText(")=mim (")
				.AddText($"{ecoRenRpt.design_sendan_tmu}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("、")
				.AddText($"{ecoRenRpt.shukyoku_design_sendan_tmin}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")=")
				.AddText($"min({ecoRenRpt.design_sendan_tmu},{ecoRenRpt.shukyoku_design_sendan_tmin})").SetFontColor(XLColor.Red).SetBold(true);

			#endregion

			#region エコレンの算出
			//        =〔{[(min(design_sendan_tmu,shukyoku_design_sendan_tmin)/min(shukyoku_QSU_B,shukyoku_Qm_B))－0.053･span_max_Pt0.23･(Fc+18)･(1-1.61･セル[A2]/BodyHeight)/(design_sendan_MQd2+0.12)]/0.85}2-(shukyoku_S_hokyou_sPw･shukyoku_S_hokyou_wft)〕/セル[D2]
			cell = targetSheet.Cell("B29");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("        =〔{[(")
				.AddText($"min({ecoRenRpt.design_sendan_tmu},{ecoRenRpt.shukyoku_design_sendan_tmin})").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/")
				.AddText($"min({ecoRenRpt.shukyoku_QSU_B},{ecoRenRpt.shukyoku_Qm_B})").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")－0.053･")
				.AddText($"{ecoRenRpt.span_max_Pt}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("0.23").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText("･(")
				.AddText($"{ecoRenRpt.Fc}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+18)･(1-1.61･")
				.AddText("セル[A2]").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/")
				.AddText($"{ecoRenRpt.BodyHeight}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")/(")
				.AddText($"{ecoRenRpt.design_sendan_MQd2}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+0.12)]/0.85}")
				.AddText("2").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText("-(")
				.AddText($"{ecoRenRpt.shukyoku_S_hokyou_sPw}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.shukyoku_S_hokyou_wft}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")〕/")
				.AddText("セル[D2]").SetFontColor(XLColor.Red).SetBold(true);

			//        =needs_danmen_shukyoku_mage/(BodyWidth*rangeC_C)
			cell = targetSheet.Cell("B30");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText($"        ={ecoRenRpt.needs_danmen_shukyoku_mage}/({ecoRenRpt.BodyWidth}*{ecoRenRpt.rangeC_C})").SetFontColor(XLColor.Red).SetBold(true);

			//sPw=(nS･aw)/(b･c) = (rangeC_Cst･(sendan_hokyo_S･stpCount))/(BodyWidth･rangeC_C)=shukyoku_S_hokyou_sPw　（C区間St=rangeC_Cst組）
			cell = targetSheet.Cell("B32");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("sPw=(nS･aw)/(b･c) = (")
				.AddText($"{ecoRenRpt.rangeC_Cst}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･(")
				.AddText($"{ecoRenRpt.sendan_hokyo_S}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.stpCount}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("))/(")
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.rangeC_C}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")=")
				.AddText($"{ecoRenRpt.shukyoku_S_hokyou_sPw}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("　（C区間St=")
				.AddText($"{ecoRenRpt.rangeC_Cst}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("組）");

			//ePa=ePw0･b･C=-0.00133･BodyWidth･rangeC_C=needs_danmen_shukyoku_mage (㎟)　　エコレンの補強必要断面積
			cell = targetSheet.Cell("B33");
			cell.Clear();
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("ePa=ePw").SetFontSize(9)
				.AddText("0").SetFontSize(9).VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("･b･C=-0.00133･").SetFontSize(9)
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true).SetFontSize(9)
				.AddText("･").SetFontSize(9)
				.AddText($"{ecoRenRpt.rangeC_C}").SetFontColor(XLColor.Red).SetBold(true).SetFontSize(9)
				.AddText("=").SetFontSize(9)
				.AddText($"{ecoRenRpt.needs_danmen_shukyoku_mage}").SetFontColor(XLColor.Red).SetBold(true).SetFontSize(9)
				.AddText(" (㎟)　　エコレンの補強必要断面積").SetFontSize(9);

			//      =needs_danmen_shukyoku_mage(㎟) ≦shukyoku_Qm_S㎟ 　エコレン shukyoku　　OK
			cell = targetSheet.Cell("B34");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("      =")
				.AddText($"{ecoRenRpt.needs_danmen_shukyoku_mage}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("(㎟) ≦")
				.AddText($"{ecoRenRpt.shukyoku_Qm_S}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("㎟ 　エコレン ")
				.AddText($"{ecoRenRpt.shukyoku}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("　　OK");
			#endregion

			#region 開孔梁のせん断耐力
			//        =0.053･span_max_Pt0.23･(Fc+18)･(1-1.61･セル[A2]/BodyHeight)/(design_sendan_MQd2+0.12)+0.85･(shukyoku_Qm_ePw･セル[D2]+shukyoku_S_hokyou_sPw･shukyoku_S_hokyou_wft)0.5
			cell = targetSheet.Cell("B38");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("        =0.053･")
				.AddText($"{ecoRenRpt.span_max_Pt}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("0.23").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
			richText
				.AddText("･(")
				.AddText($"{ecoRenRpt.Fc}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+18)･(1-1.61･")
				.AddText($"セル[A2]").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("/")
				.AddText($"{ecoRenRpt.BodyHeight}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")/(")
				.AddText($"{ecoRenRpt.design_sendan_MQd2}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+0.12)+0.85･(")
				.AddText($"{ecoRenRpt.shukyoku_Qm_ePw}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText("セル[D2]").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("+")
				.AddText($"{ecoRenRpt.shukyoku_S_hokyou_sPw}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･")
				.AddText($"{ecoRenRpt.shukyoku_S_hokyou_wft}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(")")
				.AddText("0.5").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;

			//        =shukyoku_Qm_utH(N/㎟)
			cell = targetSheet.Cell("B39");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("        =")
				.AddText($"{ecoRenRpt.shukyoku_Qm_utH}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("(N/㎟)");

			//uτH ≦ uτD　　shukyoku_Qm_utD ≦ shukyoku_Qm_utH　OK
			cell = targetSheet.Cell("B40");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("uτH ≦ uτ")
				.AddText("D").VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("　　")
				.AddText($"{ecoRenRpt.shukyoku_Qm_utD}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" ≦ ")
				.AddText($"{ecoRenRpt.shukyoku_Qm_utH}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("　OK");

			//QH = uτH･ｂ･j = shukyoku_Qm_utH･BodyWidth･span_j = shukyoku_Qm_utH･BodyWidth･span_j (kN)
			cell = targetSheet.Cell("B41");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("Q").SetFontColor(XLColor.Black)
				.AddText("H").SetFontColor(XLColor.Black).VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText(" = uτH･ｂ･j = ").SetFontColor(XLColor.Black)
				.AddText($"{ecoRenRpt.shukyoku_Qm_utH}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･").SetFontColor(XLColor.Black)
				.AddText($"{ecoRenRpt.BodyWidth}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("･").SetFontColor(XLColor.Black)
				.AddText($"{ecoRenRpt.span_j}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" = ").SetFontColor(XLColor.Black)
				.AddText($"{ecoRenRpt.shukyoku_Qm_utH}･{ecoRenRpt.BodyWidth}･{ecoRenRpt.span_j}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" (kN)").SetFontColor(XLColor.Black);

			//QD ≦ QH　　 min(design_sendan_tmu,shukyoku_design_sendan_tmin) ≦ shukyoku_Qm_utH･BodyWidth･span_j　OK
			cell = targetSheet.Cell("B42");
			richText = cell.GetRichText();
			richText.ClearText()
				.AddText("Q").SetFontColor(XLColor.Black)
				.AddText("D").SetFontColor(XLColor.Black).VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText(" ≦ ").SetFontColor(XLColor.Black)
				.AddText("Q").SetFontColor(XLColor.Black)
				.AddText("H").SetFontColor(XLColor.Black).VerticalAlignment = XLFontVerticalTextAlignmentValues.Subscript;
			richText
				.AddText("　　 ")
				.AddText($"min({ecoRenRpt.design_sendan_tmu},{ecoRenRpt.shukyoku_design_sendan_tmin})").SetFontColor(XLColor.Red).SetBold(true)
				.AddText(" ≦ ").SetFontColor(XLColor.Black)
				.AddText($"{ecoRenRpt.shukyoku_Qm_utH}･{ecoRenRpt.BodyWidth}･{ecoRenRpt.span_j}").SetFontColor(XLColor.Red).SetBold(true)
				.AddText("　OK").SetFontColor(XLColor.Black);


			#endregion

			#endregion

			//保存時の形状非表示の問題を解決する 
			//参考:https://github.com/ClosedXML/ClosedXML/issues/1252
			using (var bitmap = new Bitmap(1, 1))
			{
				using (var stream = new MemoryStream())
				{
					bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
					stream.Position = 0;
					var picture = targetSheet.AddPicture(stream)
						.MoveTo(targetSheet.Cell("BF999"));
					workbook.Save();
				}
			}
		}

	}
}

