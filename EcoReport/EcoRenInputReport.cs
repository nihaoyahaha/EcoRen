using EcoRecDatas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace EcoReport
{
    public class EcoRenInputReport:EcoRenInput
    {	
		/// <summary>
		/// 主筋径の正当性検査
		/// </summary>
		/// <param name="dm"></param>
		/// <returns></returns>
		public bool IsExistsMainDm(double dm)
		{
			if (_CTableList.DmList.Exists(d => d.Dm == dm) == false)
			{
				return false;
			}
			return true;
		}

		/// <summary>
		/// 主筋強度種別の正当性検査
		/// </summary>
		/// <param name="mainKind"></param>
		/// <returns></returns>
		public bool IsExistsMainKind(int mainKind)
		{
			if (_CTableList.KindList.Exists(d => d == mainKind) == false)
			{
				return false;
			}
			return true;
		}

		/// <summary>
		/// スターラップ径の正当性検査
		/// </summary>
		/// <param name="stpDm"></param>
		/// <returns></returns>
		public bool IsExistsStpDm(double stpDm)
		{
			if (_CTableList.DmList.Exists(d => d.Dm == stpDm) == false)
			{
				return false;
			}
			return true;
		}

		/// <summary>
		/// スターラップ強度種別の正当性検査
		/// </summary>
		/// <param name="stpKind"></param>
		/// <returns></returns>
		public bool IsExistsStpKind(int stpKind)
		{
			if (_CTableList.KindList.Exists(d => d == stpKind) == false)
			{
				return false;
			}
			return true;
		}

	}

}
