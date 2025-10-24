using Microsoft.Extensions.Configuration;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EcoReport
{
	public static class GlobalLog
	{
		private static readonly object _lock = new object();
		private static bool _isInitialized = false;

		public static void EnsureInitialized()
		{
			if (_isInitialized) return;

			lock (_lock)
			{
				if (_isInitialized) return;

				try
				{
					var config = new ConfigurationBuilder()
						.AddJsonFile("App.json", optional: false, reloadOnChange: true)
						   .Build();

					Log.Logger = new LoggerConfiguration()
						.ReadFrom.Configuration(config)
						.CreateLogger();

					_isInitialized = true;

					Log.Information($"{DateTime.Now.ToString()}:ログシステムが初期化されました");
				}
				catch (Exception ex)
				{
					_isInitialized = true;
					Log.Error($"{DateTime.Now.ToString()}ログの初期化に失敗しました: {ex.Message}");
				}
			}
		}

		public static void Information(string message, params object[] args) => Log.Information(message, args);

		public static void Warning(string message, params object[] args) => Log.Warning(message, args);

		public static void Error(string message, params object[] args) => Log.Error(message, args);

		public static void Error(Exception ex, string message, params object[] args) => Log.Error(ex, message, args);

		public static void Debug(string message, params object[] args) => Log.Debug(message, args);

		public static void Verbose(string message, params object[] args) => Log.Verbose(message, args);

		public static void Fatal(string message, params object[] args) => Log.Fatal(message, args);

		public static void Fatal(Exception ex, string message, params object[] args) => Log.Fatal(ex, message, args);
	}
}
