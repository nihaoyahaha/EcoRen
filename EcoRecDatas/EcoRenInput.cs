using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EcoRecDatas
{
	public class EcoRen
	{
		// 現場の情報を追加してください

		public List<EcoRenInput> Datas = new List<EcoRenInput>();

		// ファイル入出力の機能を追加してください。
	}


	public class EcoRenInput
	{
		#region イベント

		// 不正な値を検出した場合発生する
		public enum InputDataWarnings { mainDm = 0, stpDm, mainKind, stpKind }
		public class InputDataWarningsEventArgs : EventArgs
		{
			public InputDataWarnings inputDataWarning;
			public string strMessage;
		}
		public event EventHandler<InputDataWarningsEventArgs> InputDataWarningsEvent = null;
		#endregion

		protected TableList _CTableList = new TableList(); //テーブルのクラス

		#region 入力 --------------------------------------------------------------------

		private double _mainDm;
		private double _stpDm;
		private int _mainKind;
		private int _stpKind;

		public string strFloor = ""; //階
		public string strListName = ""; //梁符号
		public double mainDm //主筋径
		{
			get
			{
				return _mainDm;
			}
			set
			{
				_mainDm = value;
				if (_CTableList.DmList.Exists(d => d.Dm == value) == false)
				{
					if (InputDataWarningsEvent != null)
						InputDataWarningsEvent.Invoke(
							this,
							new InputDataWarningsEventArgs {
								strMessage =$"主筋径が正しくありません。{value}",
								inputDataWarning = InputDataWarnings.mainDm
							});
				}
			}
		}
		public int mainKind //主筋強度種別
		{
			get
			{
				return _mainKind;
			}
			set
			{
				_mainKind = value;
				if (_CTableList.KindList.Exists(d => d == value) == false)
				{
					if (InputDataWarningsEvent != null)
						InputDataWarningsEvent.Invoke(
							this,
							new InputDataWarningsEventArgs
							{
								strMessage = $"主筋強度が正しくありません。{value}",
								inputDataWarning = InputDataWarnings.mainKind
							});
				}
			}
		}
		public int BodyWidth; //梁幅
		public int BodyHeight; //梁背
		public double Fc; //コンクリート強度(doubleで無くていいのか？)
		public int Kaburi; //上下かぶり合計

		public int LeftHiCount; //左端上筋総本数
		public int LeftLowCount; //左端下筋総本数
		public int RightHiCount; //右端上筋総本数
		public int RightLowCount; //右端下筋総本数
		public double stpDm //スターラップ径
		{
			get
			{
				return _stpDm;
			}
			set
			{
				_stpDm = value;
				if (_CTableList.DmList.Exists(d => d.Dm == value) == false)
				{
					if (InputDataWarningsEvent != null)
						InputDataWarningsEvent.Invoke(
							this,
							new InputDataWarningsEventArgs
							{
								strMessage = $"STP径が正しくありません。{value}",
								inputDataWarning = InputDataWarnings.stpDm
							});
				}
			}
		}
		public int stpKind //スターラップ強度種別
		{
			get
			{
				return _stpKind;
			}
			set
			{
				_stpKind = value;
				if (_CTableList.KindList.Exists(d => d == value) == false)
				{
					if (InputDataWarningsEvent != null)
						InputDataWarningsEvent.Invoke(
							this,
							new InputDataWarningsEventArgs
							{
								strMessage = $"STP強度が正しくありません。{value}",
								inputDataWarning = InputDataWarnings.stpKind
							});
				}
			}
		}
		public int stpCount; //スターラップ本数
		public int stpPitch; //スターラップピッチ

		public int Uchinori; //内法 梁の長さ

		public int PAI = 150; //内径Φ
		public int EcoRenKind = 490; //エコレン強度種別

		#endregion ----------------------------------------------------------------------

		#region 計算 ##################################################################

		// 開口外径Φ
		public int out_side_PAI
		{
			get
			{
				var data = _CTableList.PAI_List.Find(d => d.in_side == PAI);
				if (data != null) return data.out_side;
				throw new EcoRenException($"製品金物に該当するものが無い。開孔={PAI}", EcoRenExceptionCodes.NonTable);
			}
		}

		// 短期エコレンwft
		public int tanki_EcoRenKind
		{
			get
			{
				var data = _CTableList.wtf_tanki_list.Find(d => d.Kind == EcoRenKind);
				if (data != null) return data.Tekioh;
				throw new EcoRenException($"wft 短期の強度変換テーブルが無い。{EcoRenKind}", EcoRenExceptionCodes.NonTable);
			}
		}

		// 枚数制限 製品の枚数がこれを超えてはならない
		public int MaisuhLimit
		{
			get
			{
				var num = (int)(BodyWidth / 100) - 1;
				return num;
			}
		}

		public double Qo
		{
			get
			{
				var C = (double)BodyWidth;
				var D = (double)BodyHeight;
				var Q = (double)Uchinori;

				var ans = (C * (D - 200) * Q * 0.024 + C * Q * 7.45 + (Q * Q) / 4 * 2 * 7.45) * (0.000001) / 2;

				return ans;
			}
		}

		#region 断面スパン ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		public double span_dc
		{
			get
			{
				var F = (double)Kaburi;
				var G = mainDm;
				var N = stpDm;

				var main_out_Dm = GetOutSideDm(mainDm);
				var stp_out_Dm = GetOutSideDm(stpDm);

				var ans = (F / 2) + stp_out_Dm + (main_out_Dm / 2);

				return ans;
			}
		}
		public double span_d
		{
			get
			{
				var D = (double)BodyHeight;
				var T = span_dc;

				return D - T;
			}
		}
		public double span_j
		{
			get
			{
				return span_d * 7 / 8;
			}
		}
		#region 左
		public int span_left_Count
		{
			get
			{
				var a = LeftHiCount + RightLowCount;
				var b = LeftLowCount + RightHiCount;
				if (a >= b) return LeftHiCount; else return LeftLowCount;
			}
		}
		public double span_left_S
		{
			get
			{
				var data = GetDmData(mainDm);
				if (data != null) return data.Danmen;
				throw new EcoRenException($"指定された径のテーブルが無い。{mainDm}mm", EcoRenExceptionCodes.NonTable);
			}
		}
		public double span_left_at
		{
			get
			{
				return span_left_Count * span_left_S;
			}
		}
		public double span_left_Pt
		{
			get
			{
				var Y = span_left_at;
				var C = (double)BodyWidth;
				var U = span_d;
				if (Y / (C * U) * 100 > 2.5) return 2.5;
				else return Y / (C * U) * 100;
			}
		}
		#endregion
		#region 右
		public int span_right_Count
		{
			get
			{
				var a = LeftHiCount + RightLowCount;
				var b = LeftLowCount + RightHiCount;
				if (a >= b) return RightLowCount; else return RightHiCount;
			}
		}
		public double span_right_S
		{
			get
			{
				var data = GetDmData(mainDm);
				if (data != null) return data.Danmen;
				throw new EcoRenException($"指定された径のテーブルが無い。{mainDm}mm", EcoRenExceptionCodes.NonTable);
			}
		}
		public double span_right_at
		{
			get
			{
				return span_right_Count * span_right_S;
			}
		}
		public double span_right_Pt
		{
			get
			{
				var AC = span_right_at;
				var C = (double)BodyWidth;
				var U = span_d;
				if (AC / (C * U) * 100 > 2.5) return 2.5;
				else return AC / (C * U) * 100;
			}
		}
		#endregion
		#region 主筋MAX
		public double span_max_at
		{
			get
			{
				return Math.Max(span_left_at, span_right_at);
			}
		}
		public double span_max_Pt
		{
			get
			{
				return Math.Max(span_left_Pt, span_right_Pt);
			}
		}
		#endregion
		#region せん断補強筋
		public double sendan_hokyo_aw
		{
			get
			{
				return sendan_hokyo_S * stpCount;
			}
		}
		public double sendan_hokyo_S
		{
			get
			{
				var data = GetDmData(stpDm);
				if (data != null) return data.Danmen;
				throw new EcoRenException($"指定された径のテーブルが無い。{stpDm}mm", EcoRenExceptionCodes.NonTable);
			}
		}
		public double sendan_hokyo_Pw
		{
			get
			{
				// IF((M*AH)/(C*O)<0.002,0.002,IF((M*AH)/(C*O)>0.012,0.012,(M*AH)/(C*O)))
				// IF(tmp<0.002,   0.002   ,IF(tmp>0.012, 0.012,tmp))

				var M = stpCount;
				var AH = sendan_hokyo_S;
				var C = (double)BodyWidth;
				var O = (double)stpPitch;

				var tmp = (M * AH) / (C * O);

				if (tmp < 0.002) return 0.002;
				if (0.012 < tmp) return 0.012;
				return tmp;
			}
		}

		public double tanki_Pw
		{
			get
			{
				return sendan_hokyo_Pw;
			}
		}
		public double chohki_Pw
		{
			get
			{
				//IF(AI>0.006,0.006,AI)
				var AI = sendan_hokyo_Pw;
				if (AI > 0.006) return 0.006; else return AI;
			}
		}
		public double shukyoku_Pw
		{
			get
			{
				return sendan_hokyo_Pw;
			}
		}
		#endregion
		#region 曲げ耐力（設計用せん断力）
		public double design_sendan_HiMy
		{
			get
			{
				// 0.9*(Y/1000)*1.1*L*(U/1000)
				var Y = span_left_at;
				var L = mainKind;
				var U = span_d;

				return 0.9 * (Y / 1000) * 1.1 * L * (U / 1000);
			}
		}
		public double design_sendan_LowMy
		{
			get
			{
				// 0.9*(AC/1000)*1.1*L*(U/1000)
				var AC = span_right_at;
				var L = mainKind;
				var U = span_d;

				return 0.9 * (AC / 1000) * 1.1 * L * (U / 1000);
			}
		}
		public double design_sendan_My
		{
			get
			{

				return design_sendan_HiMy + design_sendan_LowMy;
			}
		}
		public double design_sendan_max_My
		{
			get
			{
				return Math.Max(design_sendan_HiMy, design_sendan_LowMy);
			}
		}
		public double design_sendan_max_Qmu
		{
			get
			{
				//S + (AS * AM / (Q / 1000))
				//S+(AS*AM/(Q/1000))
				var S = Qo;
				var AS = design_sendan_a;
				var AM = design_sendan_My;
				var Q = (double)Uchinori;

				var ans = S + (AS * AM / (Q / 1000));

				return ans;
			}
		}
		public double design_sendan_a
		{
			get
			{
				return 1.1;
			}
		}
		public double design_sendan_tmu
		{
			get
			{
				//AQ/(C*V/1000)
				var AQ = design_sendan_max_Qmu;
				var C = BodyWidth;
				var V = span_j;

				var ans = AQ / (C * V / 1000);

				return ans;
			}
		}
		public double design_sendan_MQd1
		{
			get
			{
				//(MAX(AN,AO))/(AQ*(U/1000))
				var AN = design_sendan_HiMy;
				var AO = design_sendan_LowMy;
				var AQ = design_sendan_max_Qmu;
				var U = span_d;

				var ans = Math.Max(AN, AO) / (AQ * (U / 1000));
				return ans;
			}
		}
		public double design_sendan_MQd2
		{
			get
			{
				//IF(AT<1,1,IF(AT>3,3,AT))

				var AT = design_sendan_MQd1;

				if (AT < 1) return 1;
				if (3 < AT) return 3;
				return AT;
			}
		}
		#endregion
		#region 短期設計用せん断力
		public double tanki_design_sendan_a0
		{
			get
			{
				//  4/(AU+1)
				var AU = design_sendan_MQd2;
				var ans = 4 / (AU + 1);
				return ans;
			}
		}
		public double tanki_design_sendan_a
		{
			get
			{
				// IF(AV10<1,1,IF(AV10>2,2,AV10))
				var AV = tanki_design_sendan_a0;

				if (AV < 1) return 1;
				if (2 < AV) return 2;

				return AV;
			}
		}
		public double tanki_design_sendan_fs_30
		{
			get
			{
				//E/30
				var E = (double)Fc;
				return E / 30;
			}
		}
		public double tanki_design_sendan_fs_49
		{
			get
			{
				//0.49+E/100
				var E = (double)Fc;
				return 0.49 + E / 100;
			}
		}
		public double tanki_design_sendan_fs
		{
			get
			{
				// 1.5*MIN(AX,AY)
				var AX = tanki_design_sendan_fs_30;
				var AY = tanki_design_sendan_fs_49;
				return 1.5 * Math.Min(AX, AY);
			}
		}
		public double tanki_design_sendan_wft
		{
			get
			{
				var data = _CTableList.wtf_tanki_list.Find(d => d.Kind == stpKind);
				if (data != null) return data.Tekioh;
				throw new EcoRenException($"指定された強度の変換テーブルが無い。{stpKind}", EcoRenExceptionCodes.NonTable);
			}
		}
		public double tanki_design_sendan_tAS
		{
			get
			{
				//IF(P>=685,(2/3)*AW*AZ+(0.5*BA*(AJ-0.001)),(2/3)*AW*AZ+(0.5*BA*(AJ-0.002)))
				var BA = tanki_design_sendan_wft;
				var P = stpKind;
				var AW = tanki_design_sendan_a;
				var AZ = tanki_design_sendan_fs;
				var AJ = tanki_Pw;
				if (P >= 685) return (2.0 / 3) * AW * AZ + (0.5 * BA * (AJ - 0.001));
				else return (2.0 / 3) * AW * AZ + (0.5 * BA * (AJ - 0.002));

			}
		}
		public double tanki_design_sendan_QAS
		{
			get
			{
				// (C/1000)*(V/1000)*BB
				var C = (double)BodyWidth;
				var V = span_j;
				var BB = tanki_design_sendan_tAS;
				var ans = (C / 1000) * (V / 1000) * BB;
				return ans;
			}


		}
		#endregion
		#region 長期設計用せん断力
		public double chohki_design_sendan_a0
		{
			get
			{
				//  4/(AU+1)
				var AU = design_sendan_MQd2;
				var ans = 4 / (AU + 1);
				return ans;
			}
		}
		public double chohki_design_sendan_a
		{
			get
			{
				//IF(BD<1,1,IF(BD>2,2,BD10))
				var BD = chohki_design_sendan_a0;
				if (BD < 1) return 1;
				if (2 < BD) return 2;

				return BD;
			}
		}
		public double chohki_design_sendan_fs_30
		{
			get
			{
				//E/30
				var E = (double)Fc;
				return E / 30;
			}
		}
		public double chohki_design_sendan_fs_49
		{
			get
			{
				//0.49+E/100
				var E = (double)Fc;
				return 0.49 + E / 100;
			}
		}
		public double chohki_design_sendan_fs
		{
			get
			{
				var BF = chohki_design_sendan_fs_30;
				var BG = chohki_design_sendan_fs_49;
				return Math.Min(BF, BG);
			}
		}
		public double chohki_design_sendan_wft
		{
			get
			{
				return 195;
			}
		}
		public double chohki_design_sendan_tAL
		{
			get
			{
				//BE*BH
				var BE = chohki_design_sendan_a;
				var BH = chohki_design_sendan_fs;

				return BE * BH;
			}
		}
		public double chohki_design_sendan_QAL
		{
			get
			{
				// (C/1000)*(V/1000)*BJ
				var C = (double)BodyWidth;
				var V = span_j;
				var BJ = chohki_design_sendan_tAL;
				var ans = (C / 1000) * (V / 1000) * BJ;
				return ans;
			}


		}
		#endregion
		#region 終局設計用せん断力
		public double shukyoku_design_sendan_Qmin
		{
			get
			{
				// ((((0.053*(AF^0.23)*(E+18))/(AU+0.12))+(0.85*(AL*P)^0.5)))*C*V/1000
				var AF = span_max_Pt;
				var E = Fc;
				var AU = design_sendan_MQd2;
				var AL = shukyoku_Pw;
				var P = (double)stpKind;
				var C = (double)BodyWidth;
				var V = span_j;

				var ans = ((((0.053 * Math.Pow(AF, 0.23) * (E + 18)) / (AU + 0.12)) + (0.85 * Math.Pow((AL * P), 0.5)))) * C * V / 1000;

				return ans;
			}
		}
		public double shukyoku_design_sendan_RightSide
		{
			get
			{
				// ((((0.053*(AF^0.23)*(E+18))/(AU+0.12))))*C*V/1000
				var AF = span_max_Pt;
				var E = Fc;
				var AU = design_sendan_MQd2;
				var P = (double)stpKind;
				var C = (double)BodyWidth;
				var V = span_j;

				var ans = ((((0.053 * Math.Pow(AF, 0.23) * (E + 18)) / (AU + 0.12)))) * C * V / 1000;

				return ans;
			}
		}
		public double shukyoku_design_sendan_LeftSide
		{
			get
			{
				// (0.85*(AL*P)^0.5)*C*V/1000
				var AL = shukyoku_Pw;
				var P = (double)stpKind;
				var C = (double)BodyWidth;
				var V = span_j;

				var ans = (0.85 * Math.Pow((AL * P), 0.5)) * C * V / 1000;

				return ans;
			}
		}
		public double shukyoku_design_sendan_tmin
		{
			get
			{
				// ((0.053*(AF^0.23)*(E+18))/(AU+0.12))+(0.85*(AL*P)^0.5)
				var AF = span_max_Pt;
				var E = Fc;
				var AU = design_sendan_MQd2;
				var AL = shukyoku_Pw;
				var P = (double)stpKind;

				var ans = ((0.053 * Math.Pow(AF, 0.23) * (E + 18)) / (AU + 0.12)) + (0.85 * Math.Pow((AL * P), 0.5));

				return ans;
			}
		}
		public double shukyoku_design_sendan_tmean
		{
			get
			{
				// (((0.068*(AF^0.23)*(E+18))/(AU+0.12))+(0.85*(AL*P)^0.5)/1.1)
				var AF = span_max_Pt;
				var E = Fc;
				var AU = design_sendan_MQd2;
				var AL = shukyoku_Pw;
				var P = (double)stpKind;
				var C = (double)BodyWidth;
				var V = span_j;

				var ans = ((0.068 * Math.Pow(AF, 0.23) * (E + 18)) / (AU + 0.12))
					+ (0.85 * Math.Pow((AL * P), 0.5)) / 1.1;

				return ans;
			}
		}
		#endregion
		#region 開孔補強部分の計算(C区間)
		public double rangeC_C
		{
			get
			{
				// D-2*T
				var D = BodyHeight;
				var T = span_dc;
				return D - 2 * T;
			}
		}
		public double rangeC_C1
		{
			get
			{
				// D/2-T
				var D = BodyHeight;
				var T = span_dc;
				return D / 2 - T;
			}
		}
		public double rangeC_C2
		{
			get
			{
				// D/2-T
				var D = BodyHeight;
				var T = span_dc;
				return D / 2 - T;
			}
		}
		public int rangeC_Cst
		{
			get
			{
				return rangeC_C1st + rangeC_C2st;
			}
		}
		public int rangeC_C1st
		{
			get
			{
				///IF($B$2<150,INT(1+(BR-($B$2/2)-50)/O),INT(2+(BR-($B$2/2)-100)/O)

				var PI = out_side_PAI;
				var BR = rangeC_C1;
				var O = stpPitch;

				if (PI < 150) return (int)(1 + (BR - (PI / 2) - 50) / O);
				else return (int)(2 + (BR - (PI / 2) - 100) / O);

			}
		}
		public int rangeC_C2st
		{
			get
			{
				///IF($B$2<150,INT(1+(BR-($B$2/2)-50)/O),INT(2+(BR-($B$2/2)-100)/O)

				var PI = out_side_PAI;
				var BR = rangeC_C1;
				var O = stpPitch;

				if (PI < 150) return (int)(1 + (BR - (PI / 2) - 50) / O);
				else return (int)(2 + (BR - (PI / 2) - 100) / O);

			}
		}
		#endregion
		#region 開孔補強部分の計算(S筋・短期)
		public double tanki_S_hokyou_sPw00
		{
			get
			{
				// BX*BZ
				var BX = tanki_S_hokyou_sPw;
				var BZ = tanki_S_hokyou_wft;

				return BX * BZ;
			}
		}
		public double tanki_S_hokyou_sPw
		{
			get
			{
				// BT*M*AH/(C*BQ)
				var BT = rangeC_Cst;
				var M = (double)stpCount;
				var AH = sendan_hokyo_S;
				var C = (double)BodyWidth;
				var BQ = rangeC_C;
				var ans = BT * M * AH / (C * BQ);

				return ans;
			}
		}
		public double tanki_S_hokyou_sPw12
		{
			get
			{
				// (IF(BX>0.012,0.012,BX))
				var BX = tanki_S_hokyou_sPw;

				if (BX > 0.012) return 0.012;
				return BX;
			}
		}
		public double tanki_S_hokyou_wft
		{
			get
			{
				return tanki_design_sendan_wft;
			}
		}
		#endregion
		#region 開孔補強部分の計算(S筋・長期)
		public double chohki_S_hokyou_sPw00
		{
			get
			{
				// BX*BZ
				var BX = chohki_S_hokyou_sPw;
				var BZ = chohki_S_hokyou_wft;

				return BX * BZ;
			}
		}
		public double chohki_S_hokyou_sPw
		{
			get
			{
				// BT*M*AH/(C*BQ)
				var BT = rangeC_Cst;
				var M = (double)stpCount;
				var AH = sendan_hokyo_S;
				var C = (double)BodyWidth;
				var BQ = rangeC_C;
				var ans = BT * M * AH / (C * BQ);

				return ans;
			}
		}
		public double chohki_S_hokyou_sPw12
		{
			get
			{
				// (IF(CB>0.006,0.006,CB))
				var CB = tanki_S_hokyou_sPw;

				if (CB > 0.006) return 0.006;
				return CB;
			}
		}
		public double chohki_S_hokyou_wft
		{
			get
			{
				return chohki_design_sendan_wft;
			}
		}
		#endregion
		#region 開孔補強部分の計算(S筋・終局)
		public double shukyoku_S_hokyou_sPw00
		{
			get
			{
				// CF*CG
				var CF = shukyoku_S_hokyou_sPw;
				var CG = shukyoku_S_hokyou_wft;

				return CF * CG;
			}
		}
		public double shukyoku_S_hokyou_sPw
		{
			get
			{
				// BT*M*AH/(C*BQ)
				var BT = rangeC_Cst;
				var M = (double)stpCount;
				var AH = sendan_hokyo_S;
				var C = (double)BodyWidth;
				var BQ = rangeC_C;
				var ans = BT * M * AH / (C * BQ);

				return ans;
			}
		}
		public double shukyoku_S_hokyou_wft
		{
			get
			{
				return stpKind;
			}
		}
		#endregion
		#region 検証(短期)
		public double tanki_test_PA
		{
			get
			{
				// IF(P>=685,(BB-AW*AZ*(1-(1.61*$B$2/D))-0.5*BA*(BX-0.001))/(0.5*$CH$3),(BB-AW*AZ*(1-(1.61*$B$2/D))-0.5*BA*(BX-0.002))/(0.5*$CH$3))

				var B2 = out_side_PAI;
				var CH3 = tanki_EcoRenKind;
				var P = stpKind;
				var BB = tanki_design_sendan_tAS;
				var AW = tanki_design_sendan_a;
				var AZ = tanki_design_sendan_fs;
				var BA = tanki_design_sendan_wft;
				var BX = tanki_S_hokyou_sPw;
				var D = BodyHeight;

				if(P >= 685)
					return (BB - AW * AZ * (1 - (1.61 *B2 / D))-0.5 * BA * (BX - 0.001))/ (0.5 *CH3);
				else 
					return (BB - AW * AZ * (1 - (1.61 *B2 / D))-0.5 * BA * (BX - 0.002))/ (0.5 *CH3);


			}
		}
		public double tanki_test_aa
		{
			get
			{
				// CH*C*BQ
				var CH = tanki_test_PA;
				var C = BodyWidth;
				var BQ = rangeC_C;

				var ans = CH * C * BQ;

				return ans;
			}
		}
		#endregion
		#region 検証(長期)
		public double chohki_test_PA
		{
			get
			{
				// (BJ-BE*BH*(1-(1.61*B2/D))-0.5*BI*(CC-0.002))/(0.5*BI)
				// (BJ-BE*BH*(1-(1.61*B2/D))-0.5*BI*(CC-0.002))/(0.5*BI)
				var B2 = out_side_PAI;
				var CH3 = tanki_EcoRenKind;

				var BJ = chohki_design_sendan_tAL;
				var BE = chohki_design_sendan_a;
				var BH = chohki_design_sendan_fs;
				var BI = chohki_design_sendan_wft;
				var CC = chohki_S_hokyou_sPw12;
				var D = (double)BodyHeight;

				var ans = (BJ - BE * BH * (1 - (1.61 * B2 / D)) - 0.5 * BI * (CC - 0.002)) / (0.5 * BI);
				return ans;

			}
		}
		public double chohki_test_aa
		{
			get
			{
				// C*BQ*CL
				var C = BodyWidth;
				var BQ = rangeC_C;
				var CL = chohki_test_PA;

				var ans = CL * C * BQ;

				return ans;
			}
		}
		#endregion
		#region 設計用せん断力　QSU　終局の場合
		public double shukyoku_QSU_B
		{
			get { return 1; }
		}
		public double shukyoku_QSU_utD
		{
			get
			{
				// BO
				var BO = shukyoku_design_sendan_tmin;
				return BO;
			}
		}
		public double shukyoku_QSU_aa
		{
			get
			{
				// DI
				var DI = needs_danmen_shukyoku;
				return DI;
			}
		}
		public double shukyoku_QSU_S
		{
			get
			{
				var EQ = tmpQSU;
				return EQ;
			}
		}
		public double shukyoku_QSU_ePw
		{
			get
			{
				//CQ/(C*BQ)
				var CQ = shukyoku_QSU_S;
				var C = BodyWidth;
				var BQ = rangeC_C;
				var ans = CQ / (C * BQ);

				return ans;
			}
		}
		public double shukyoku_QSU_utH
		{
			get
			{
				// ((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12))))+(0.85*(CE+D2*CR)^0.5)

				var B2 = out_side_PAI;
				var D2 = EcoRenKind;

				var AF = span_max_Pt;
				var E = Fc;
				var D = BodyHeight;
				var CE = shukyoku_S_hokyou_sPw00;
				var AU = design_sendan_MQd2;
				var CR = shukyoku_QSU_ePw;

				var arg = ((((0.053 * Math.Pow(AF, 0.23) * (E + 18) * (1 - (1.61 * B2 / D))) / (AU + 0.12)))) + (0.85 * Math.Pow((CE + D2 * CR), 0.5));

				return arg;
			}
		}
		public double shukyoku_QSU_utHutD
		{
			get
			{
				// CS/CO

				var CS = shukyoku_QSU_utH;
				var CO = shukyoku_QSU_utD;
				return CS / CO;
			}
		}
		public double shukyoku_QSU_SPw
		{
			get
			{
				// ((((CO/CN-CS)/0.85)^2))
				var CO = shukyoku_QSU_utD;
				var CN = shukyoku_QSU_B;
				var CS = shukyoku_QSU_utH;

				var ans = Math.Pow(((CO / CN - CS) / 0.85), 2);

				return ans;
			}
		}
		#endregion

		#region 設計用せん断力　Qm　両端曲げ降伏の場合
		public double shukyoku_Qm_utD
		{
			get
			{
				// AR
				var AR = design_sendan_tmu;
				return AR;
			}
		}
		public double shukyoku_Qm_B
		{
			get { return 1; }
		}
		public double shukyoku_Qm_aa
		{
			get
			{
				var DJ = needs_danmen_shukyoku_mage;

				return DJ;
			}
		}
		public double shukyoku_Qm_S
		{
			get
			{
				return tmpQm;
			}
		}
		public double shukyoku_Qm_ePw
		{
			get
			{
				// CY/(C*BQ)
				var BQ = rangeC_C;
				var CY = shukyoku_Qm_S;
				var C = BodyWidth;

				var ans = CY / (C * BQ);

				return ans;
			}
		}
		public double shukyoku_Qm_utH
		{
			get
			{
				// ((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12))))+(0.85*(CE+D2*CZ)^0.5)

				var B2 = out_side_PAI;
				var D2 = EcoRenKind;
				var AF = span_max_Pt;
				var E = Fc;
				var D = BodyHeight;
				var AU = design_sendan_MQd2;
				var CE = shukyoku_S_hokyou_sPw00;
				var CZ = shukyoku_Qm_ePw;

				var ans = ((((0.053 * Math.Pow(AF, 0.23) * (E + 18) * (1 - (1.61 * B2 / D))) / (AU + 0.12)))) + (0.85 * Math.Pow((CE + D2 * CZ), 0.5));

				return ans;
			}
		}
		public double shukyoku_Qm_utHutD
		{
			get
			{
				return shukyoku_Qm_utH / shukyoku_Qm_utD;
			}
		}
		#endregion


		#region 使用金物比較

		public double needs_danmen_tanki
		{
			get
			{
				// IF(P>=685,((BB-AW*AZ*(1-(1.61*B2/D))-0.5*BA*(BX-0.001))/(0.5*CH3))*C*BQ,((BB-AW*AZ*(1-(1.61*B2/D))-0.5*BA*(BX-0.002))/(0.5*CH3))*C*BQ)

				var CH3 = tanki_EcoRenKind;
				var B2 = out_side_PAI;
				var P = stpKind;
				var BB = tanki_design_sendan_tAS;
				var AW = tanki_design_sendan_a;
				var AZ = tanki_design_sendan_fs;
				var D = BodyHeight;
				var BA = tanki_design_sendan_wft;
				var BX = tanki_S_hokyou_sPw;
				var C = BodyWidth;
				var BQ = rangeC_C;

				if (P >= 685) return ((BB - AW * AZ * (1 - (1.61 * B2 / D)) - 0.5 * BA * (BX - 0.001)) / (0.5 * CH3)) * C * BQ;
				else return ((BB - AW * AZ * (1 - (1.61 * B2 / D)) - 0.5 * BA * (BX - 0.002)) / (0.5 * CH3)) * C * BQ;
			}
		}
		public double needs_danmen_chohki
		{
			get
			{
				// ((BJ-BE*BH*(1-(1.61*B2/D))-0.5*BI*(CC-0.002))/(0.5*CD))*C*BQ

				var CH3 = tanki_EcoRenKind;
				var B2 = out_side_PAI;
				var D = BodyHeight;
				var C = BodyWidth;
				var BQ = rangeC_C;
				var BJ = chohki_design_sendan_tAL;
				var BE = chohki_design_sendan_a;
				var BH = chohki_design_sendan_fs;
				var BI = chohki_design_sendan_wft;
				var CC = chohki_S_hokyou_sPw12;
				var CD = chohki_S_hokyou_wft;

				var ans = ((BJ - BE * BH * (1 - (1.61 * B2 / D)) - 0.5 * BI * (CC - 0.002)) / (0.5 * CD)) * C * BQ;

				return ans;
			}
		}
		public double needs_danmen_shukyoku
		{
			get
			{
				// IF(A2<150,(((((CO/CN-((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12)))))/0.85)^2))-((INT(1+(BR-(B2/2)-50)/O)*2)*M*AH/(C*BQ)*CG))/D2*C*BQ,(((((CO/CN-((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12)))))/0.85)^2))-((INT(2+(BR-(B2/2)-100)/O)*2)*M*AH/(C*BQ)*CG))/D2*C*BQ)

				var B2 = out_side_PAI;
				var D2 = EcoRenKind;

				var A2 = PAI;
				var CO = shukyoku_QSU_utD;
				var CN = shukyoku_QSU_B;
				var AF = span_max_Pt;
				var E = Fc;
				var D = BodyHeight;
				var AU = design_sendan_MQd2;
				var BR = rangeC_C1;
				var M = stpCount;
				var AH = sendan_hokyo_S;
				var BQ = rangeC_C;
				var C = BodyWidth;
				var CG = shukyoku_S_hokyou_wft;
				var O = stpPitch;

				if (A2 < 150) return ((Math.Pow(((CO / CN - ((((0.053 * Math.Pow(AF, 0.23) * (E + 18) * (1 - (1.61 * B2 / D))) / (AU + 0.12))))) / 0.85), 2)) - (((int)(1 + (BR - (B2 / 2) - 50) / O) * 2) * M * AH / (C * BQ) * CG)) / D2 * C * BQ;
				else return ((Math.Pow(((CO / CN - ((((0.053 * Math.Pow(AF, 0.23) * (E + 18) * (1 - (1.61 * B2 / D))) / (AU + 0.12))))) / 0.85), 2)) - (((int)(2 + (BR - (B2 / 2) - 100) / O) * 2) * M * AH / (C * BQ) * CG)) / D2 * C * BQ;
			}
		}
		public double needs_danmen_shukyoku_mage
		{
			get
			{
				// IF(A2<150,(((((((CV/CN)-((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12)))))/0.85)^2))-((INT(1+(BR-(B2/2)-50)/O)*2)*M*AH/(C*BQ)*CG))/D2)*C10*BQ,(((((((CV/CN)-((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12)))))/0.85)^2))-((INT(2+(BR-(B2/2)-100)/O)*2)*M*AH/(C*BQ)*CG))/D2)*C*BQ)

				// IF(A2<150,
				// (((((((CV/CN)-((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12)))))/0.85)^2))-((INT(1+(BR-(B2/2)-50)/O)*2)*M*AH/(C*BQ)*CG))/D2)*C*BQ,
				// (((((((CV/CN)-((((0.053*(AF^0.23)*(E+18)*(1-(1.61*B2/D)))/(AU+0.12)))))/0.85)^2))-((INT(2+(BR-(B2/2)-100)/O)*2)*M*AH/(C*BQ)*CG))/D2)*C*BQ)


				var B2 = out_side_PAI;
				var D2 = EcoRenKind;
				var A2 = PAI;
				var CV = shukyoku_Qm_utD;
				var CN = shukyoku_QSU_B;
				var AF = span_max_Pt;
				var E = Fc;
				var D = BodyHeight;
				var AU = design_sendan_MQd2;
				var BR = rangeC_C1;
				var M = stpCount;
				var AH = sendan_hokyo_S;
				var BQ = rangeC_C;
				var C = BodyWidth;
				var CG = shukyoku_S_hokyou_wft;
				var O = stpPitch;

				if (A2 < 150)
					return (((    Math.Pow((((CV / CN) - ((((0.053 * Math.Pow(AF , 0.23) * (E + 18) * (1 - (1.61 * B2 / D))) / (AU + 0.12))))) / 0.85) , 2)) - (((int)(1 + (BR - (B2 / 2) - 50) / O) * 2) * M * AH / (C * BQ) * CG)) / D2) * C * BQ;
				else
					return (((Math.Pow((((CV / CN) - ((((0.053 * Math.Pow(AF , 0.23) * (E + 18) * (1 - (1.61 * B2 / D))) / (AU + 0.12))))) / 0.85) , 2)) - (((int)(2 + (BR - (B2 / 2) - 100) / O) * 2) * M * AH / (C * BQ) * CG)) / D2) * C * BQ;
			}
		}
		public double needs_danmen_shukyoku_min
		{
			get
			{
				return Math.Min(needs_danmen_shukyoku, needs_danmen_shukyoku_mage);
			}
		}

		public TableList.KanamonoData C_tanki
		{
			get
			{
				// IF(D / 3 <A2, "OverPI", IF(A2 >= 150, VLOOKUP(Entrysheet!EH12, db!$P$4:$Q$100, 2, FALSE), VLOOKUP(Entrysheet!EH12, db!$Z$4:$AA$100, 2, FALSE)))

				var _CTableList = new TableList();

				var A2 = PAI;
				var D = BodyHeight;

				var S = needs_danmen_tanki; //面積のS　セルでは無いので注意

				if ((D / 3) < A2) throw new EcoRenException($"製品外径が梁背/3を超えた. 梁背={D} 開孔製品外径={A2}mm", EcoRenExceptionCodes.LargeHole);

				if (A2 >= 150)
				{
					var Datas = _CTableList.Kanamono150U.FindAll(d => d.SoDanmenseki >= S);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();
						return ans;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150L.FindAll(d => d.SoDanmenseki >= S);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();
						return ans;
					}
				}
				throw new EcoRenException($"tmpQm_L:製品金物に該当するものが無い。aa={S}", EcoRenExceptionCodes.NonTable);
			}
		}
		public TableList.KanamonoData C_chohki
		{
			get
			{
				// IF(D / 3 <A2, "OverPI", IF(A2 >= 150, VLOOKUP(Entrysheet!EH12, db!$P$4:$Q$100, 2, FALSE), VLOOKUP(Entrysheet!EH12, db!$Z$4:$AA$100, 2, FALSE)))

				var _CTableList = new TableList();

				var A2 = PAI;
				var D = BodyHeight;

				var S = needs_danmen_chohki;

				if ((D / 3) < A2) throw new EcoRenException($"製品外径が梁背/3を超えた. 梁背={D} 開孔製品外径={A2}mm", EcoRenExceptionCodes.LargeHole);

				if (A2 >= 150)
				{
					var Datas = _CTableList.Kanamono150U.FindAll(d => d.SoDanmenseki >= S);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();
						return ans;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150L.FindAll(d => d.SoDanmenseki >= S);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();
						return ans;
					}
				}
				throw new EcoRenException($"tmpQm_L:製品金物に該当するものが無い。aa={S}", EcoRenExceptionCodes.NonTable);
			}
		}
		public TableList.KanamonoData C_shukyoku
		{
			get
			{
				// IF(D / 3 <A2, "OverPI", IF(A2 >= 150, VLOOKUP(Entrysheet!EH12, db!$P$4:$Q$100, 2, FALSE), VLOOKUP(Entrysheet!EH12, db!$Z$4:$AA$100, 2, FALSE)))

				var _CTableList = new TableList();

				var A2 = PAI;
				var D = BodyHeight;

				var S = needs_danmen_shukyoku_min;

				if ((D / 3) < A2) throw new EcoRenException($"製品外径が梁背/3を超えた. 梁背={D} 開孔製品外径={A2}mm", EcoRenExceptionCodes.LargeHole);
				if (A2 >= 150)
				{
					var Datas = _CTableList.Kanamono150U.FindAll(d => d.SoDanmenseki >= S);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();
						return ans;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150L.FindAll(d => d.SoDanmenseki >= S);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();
						return ans;
					}
				}
				throw new EcoRenException($"tmpQm_L:製品金物に該当するものが無い。aa={S}", EcoRenExceptionCodes.NonTable);
			}
		}

		public string tanki
		{
			get
			{
				var data = C_tanki;
				var shube = EcoRecShubetsu();

				return $"{data.count}-{shube}{data.Dm}";
			}
		}
		public string chohki
		{
			get
			{
				var data = C_chohki;
				var shube = EcoRecShubetsu();

				return $"{data.count}-{shube}{data.Dm}";
			}
		}
		public string shukyoku
		{
			get
			{
				var data = C_shukyoku;
				var shube = EcoRecShubetsu();

				return $"{data.count}-{shube}{data.Dm}";
			}
		}



		#endregion

		#region temporaly
		public double tmpQSU_U
		{
			get
			{
				// =IF(CP<=0,db!$R$4,MINIFS(db!$P$4:$P$100,db!$P$4:$P$100,">="&CP10))
				var CP = shukyoku_QSU_aa;
				if (CP <= 0)
				{
					var Datas = _CTableList.Kanamono150U.ToList();
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150U.FindAll(d => d.SoDanmenseki >= CP);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				throw new EcoRenException($"tmpQSU_U:製品金物に該当するものが無い。aa={CP}", EcoRenExceptionCodes.NonTable);
			}
		}
		public double tmpQSU_L
		{
			get
			{
				// =IF(CP<=0,db!$R$4,MINIFS(db!$P$4:$P$100,db!$P$4:$P$100,">="&CP10))
				var CP = shukyoku_QSU_aa;
				if (CP <= 0)
				{
					var Datas = _CTableList.Kanamono150L.ToList();
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150L.FindAll(d => d.SoDanmenseki >= CP);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				throw new EcoRenException($"tmpQSU_L:製品金物に該当するものが無い。aa={CP}", EcoRenExceptionCodes.NonTable);
			}
		}
		public double tmpQSU
		{
			get
			{
				if (PAI < 150) return tmpQSU_L;
				else return tmpQSU_U;
			}
		}
		public double tmpQm_U
		{
			get
			{
				// IF(CX10<=0,db!$R$4,MINIFS(db!$P$4:$P$100,db!$P$4:$P$100,">="&CX10))
				var CX = shukyoku_Qm_aa;
				if (CX <= 0)
				{
					var Datas = _CTableList.Kanamono150U.ToList();
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150U.FindAll(d => d.SoDanmenseki >= CX);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				throw new EcoRenException($"tmpQm_U:製品金物に該当するものが無い。aa={CX}", EcoRenExceptionCodes.NonTable);
			}
		}
		public double tmpQm_L
		{
			get
			{
				// IF(CX10<=0,db!$AB$4,MINIFS(db!$Z$4:$Z$100,db!$Z$4:$Z$100,">="&CX10))
				var CX = shukyoku_Qm_aa;
				if (CX <= 0)
				{
					var Datas = _CTableList.Kanamono150L.ToList();
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				else
				{
					var Datas = _CTableList.Kanamono150L.FindAll(d => d.SoDanmenseki >= CX);
					if (Datas.Count != 0)
					{
						Datas.Sort((a, b) => (int)(a.SoDanmenseki - b.SoDanmenseki));
						var ans = Datas.First();

						return ans.SoDanmenseki;
					}
				}
				throw new EcoRenException($"tmpQm_L:製品金物に該当するものが無い。aa={CX}", EcoRenExceptionCodes.NonTable);
			}
		}
		public double tmpQm
		{
			get
			{
				if (PAI < 150) return tmpQm_L;
				else return tmpQm_U;
			}
		}
		#endregion



		#endregion +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		// DmListデータを得る
		private TableList.DmData GetDmData(double Dm)
		{
			var data = _CTableList.DmList.Find(d => d.Dm == Dm);
			if (data != null) return data;
			return null;
		}

		// 最外径を得る
		private double GetOutSideDm(double Dm)
		{
			var data = _CTableList.DmList.Find(d => d.Dm == Dm);
			if (data != null) return data.MaxDm;
			throw new EcoRenException($"指定された径のテーブルが無い。{Dm}mm", EcoRenExceptionCodes.NonTable);
		}

		private string EcoRecShubetsu()
		{
			switch (EcoRenKind)
			{
				default: return "D";
				case 685: return "S";
				case 785: return "S";
				case 1275: return "U";
			}
		}

		#endregion 計算終わり ############################
	}

	/// <summary>
	/// テーブルのクラス
	/// </summary>
	public class TableList
	{
		#region Dm
		public List<DmData> DmList = new List<DmData> {
			new DmData { Dm = 6, Danmen = 31.67, DmCategory = DmCategorys.EcoRen },
			new DmData { Dm = 7.1, Danmen = 40, DmCategory = DmCategorys.Maker },
			new DmData { Dm = 8, Danmen = 49.51, DmCategory = DmCategorys.EcoRen },
			new DmData { Dm = 10, Danmen = 71.33, MaxDm = 11 },
			new DmData { Dm = 10.7, Danmen = 90, DmCategory = DmCategorys.Maker },
			new DmData { Dm = 11.8, Danmen = 110.1, DmCategory = DmCategorys.Maker },
			new DmData { Dm = 12.6, Danmen = 125, DmCategory = DmCategorys.Maker },
			new DmData { Dm = 13, Danmen = 126.7, MaxDm = 14 },
			new DmData { Dm = 16, Danmen = 198.6, MaxDm = 18 },
			new DmData { Dm = 19, Danmen = 286.5, MaxDm = 22 },
			new DmData { Dm = 22, Danmen = 387.1, MaxDm = 26 },
			new DmData { Dm = 25, Danmen = 506.7, MaxDm = 29 },
			new DmData { Dm = 29, Danmen = 642.4, MaxDm = 33 },
			new DmData { Dm = 32, Danmen = 794.2, MaxDm = 37 },
			new DmData { Dm = 35, Danmen = 956.6, MaxDm = 40 },
			new DmData { Dm = 38, Danmen = 1140, MaxDm = 43 },
			new DmData { Dm = 41, Danmen = 1340, MaxDm = 47 },
			new DmData { Dm = 51, Danmen = -1, MaxDm = 58 },
		};

		public enum DmCategorys { Zairai = 10, Maker = 5, EcoRen = 1 }

		public class DmData
		{
			public double Dm; //公称
			public double Danmen = -1; //断面積
			public double MaxDm = -1; //最外径

			public DmCategorys DmCategory = DmCategorys.Zairai;
		}

		#endregion

		#region 種別
		public List<int> KindList = new List<int> { 295, 345, 390, 490, 590, 685, 785, 1275 };

		public List<wtf_tanki_Data> wtf_tanki_list = new List<wtf_tanki_Data>
		{
			new wtf_tanki_Data { Kind = 295, Tekioh = 295 },
			new wtf_tanki_Data { Kind = 345, Tekioh = 345 },
			new wtf_tanki_Data { Kind = 390, Tekioh = 390 },
			new wtf_tanki_Data { Kind = 490, Tekioh = 390 },
			new wtf_tanki_Data { Kind = 685, Tekioh = 590 },
			new wtf_tanki_Data { Kind = 785, Tekioh = 590 },
			new wtf_tanki_Data { Kind = 1275, Tekioh = 590 }
		};

		public class wtf_tanki_Data
		{
			public int Kind; //種別
			public int Tekioh; //適応値
		}
		#endregion

		#region 開口

		public List<PAI_Data> PAI_List = new List<PAI_Data>
		{
			new PAI_Data { in_side = 100, out_side = 115 },
			new PAI_Data { in_side = 125, out_side = 141 },
			new PAI_Data { in_side = 150, out_side = 166 },
			new PAI_Data { in_side = 175, out_side = 191 },
			new PAI_Data { in_side = 200, out_side = 216 },
			new PAI_Data { in_side = 250, out_side = 270 },
			new PAI_Data { in_side = 300, out_side = 320 },
			new PAI_Data { in_side = 350, out_side = 370 }
		};

		public class PAI_Data
		{
			public int in_side;
			public int out_side;
		}

		#endregion


		#region 金物情報

		/// <summary>
		/// 金物情報(150π以上)
		/// </summary>
		public List<KanamonoData> Kanamono150U = new List<KanamonoData>
		{
			new KanamonoData { Dm = 6, count = 2 },
			new KanamonoData { Dm = 8, count = 2 },
			new KanamonoData { Dm = 10, count = 2 },
			new KanamonoData { Dm = 13, count = 2 },
			new KanamonoData { Dm = 16, count = 2 },
			new KanamonoData { Dm = 16, count = 3 },
			new KanamonoData { Dm = 16, count = 4 },
			new KanamonoData { Dm = 16, count = 5 },
			new KanamonoData { Dm = 16, count = 6 },
			new KanamonoData { Dm = 16, count = 7 },
			new KanamonoData { Dm = 16, count = 8 },
			new KanamonoData { Dm = 16, count = 9 },
			new KanamonoData { Dm = 16, count = 10 }
		};

		/// <summary>
		/// 金物情報(150π未満)
		/// </summary>
		public List<KanamonoData> Kanamono150L = new List<KanamonoData>
		{
			new KanamonoData { Dm = 6, count = 2 },
			new KanamonoData { Dm = 8, count = 2 },
			new KanamonoData { Dm = 10, count = 2 },
			new KanamonoData { Dm = 13, count = 2 },
			new KanamonoData { Dm = 13, count = 2 },
			new KanamonoData { Dm = 13, count = 3 },
			new KanamonoData { Dm = 13, count = 4 },
			new KanamonoData { Dm = 13, count = 5 },
			new KanamonoData { Dm = 13, count = 6 },
			new KanamonoData { Dm = 13, count = 7 },
			new KanamonoData { Dm = 13, count = 8 },
			new KanamonoData { Dm = 13, count = 9 },
			new KanamonoData { Dm = 13, count = 10 }
		};

		public class KanamonoData
		{
			public double Dm;
			public int count;

			public double SoDanmenseki
			{
				get
				{
					var db = new TableList();
					var data = db.DmList.Find(d => d.Dm == Dm);
					if (data != null)
					{
						var S = data.Danmen;

						return count * 2 * S * Math.Sqrt(2);
					}
					else
					{
						return 0;
					}
				}
			}
		}

		#endregion
	}

	/// <summary>
	/// 例外の種類
	/// </summary>
	public enum EcoRenExceptionCodes { NonTable = 0, LargeHole = 1 }

	/// <summary>
	/// 製品として出せないときに発生する例外
	/// </summary>
	public class EcoRenException : Exception
	{
		public string strMessage;
		public EcoRenExceptionCodes Code;

		public EcoRenException(string msg, EcoRenExceptionCodes code)
		{
			strMessage = $"{msg}"; Code = code;
		}
	}
}
