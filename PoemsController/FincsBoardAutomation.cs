using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace PoemsController
{
	/// <summary>
	/// Fincs (NOBU塾) のチャット画面から
	/// 「ドル円→…」シグナルを取得するクラス。
	/// </summary>
	public class FincsBoardAutomation
	{
		private AutomationElement _windowRoot;
		private AutomationElement _documentRoot;

		private static readonly string[] TitleKeywords =
		{
			"NOBU塾",
			"Fincs"
		};

		/// <summary>
		/// ウィンドウと Document を特定する。
		/// </summary>
		private bool AttachInternal(StringBuilder log)
		{
			_windowRoot = null;
			_documentRoot = null;

			var desktop = AutomationElement.RootElement;
			if (desktop == null)
			{
				log?.AppendLine("RootElement が取得できませんでした。");
				return false;
			}

			AutomationElementCollection wins;
			try
			{
				wins = desktop.FindAll(TreeScope.Children, Condition.TrueCondition);
			}
			catch (Exception ex)
			{
				log?.AppendLine("ウィンドウ列挙中に例外: " + ex.Message);
				return false;
			}

			log?.AppendLine($"トップレベルウィンドウ数: {wins.Count}");

			foreach (AutomationElement w in wins)
			{
				string winName;
				int pid;
				try
				{
					winName = w.Current.Name ?? string.Empty;
					pid = w.Current.ProcessId;
				}
				catch
				{
					continue;
				}

				string procName;
				try
				{
					procName = Process.GetProcessById(pid).ProcessName;
				}
				catch
				{
					continue;
				}

				log?.AppendLine($"[Window] Proc={procName}, Name=\"{winName}\"");

				// Chrome / Edge だけ対象
				if (!procName.Equals("chrome", StringComparison.OrdinalIgnoreCase) &&
					!procName.Equals("msedge", StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}

				AutomationElementCollection docs;
				try
				{
					docs = w.FindAll(
						TreeScope.Descendants,
						new PropertyCondition(
							AutomationElement.ControlTypeProperty,
							ControlType.Document));
				}
				catch
				{
					continue;
				}

				log?.AppendLine($"  Document 数: {docs.Count}");

				foreach (AutomationElement d in docs)
				{
					string dname;
					try
					{
						dname = d.Current.Name ?? string.Empty;
					}
					catch
					{
						continue;
					}

					log?.AppendLine($"    [Document] \"{dname}\"");

					if (string.IsNullOrWhiteSpace(dname))
						continue;

					bool match = TitleKeywords.Any(
						kw => dname.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0);

					if (!match) continue;

					_windowRoot = w;
					_documentRoot = d;
					log?.AppendLine("    -> この Document を採用しました。");
					return true;
				}
			}

			log?.AppendLine("Fincs/NOBU塾 の Document が見つかりませんでした。");
			return false;
		}

		/// <summary>
		/// 通常用: デバッグログなしで最新シグナルだけ取得。
		/// </summary>
		public string? GetLatestSignal()
		{
			return ScanInternal(null);
		}

		/// <summary>
		/// デバッグ用: 探索ログも返す。
		/// </summary>
		public string? ScanDebug(out string debugLog)
		{
			var sb = new StringBuilder();
			var result = ScanInternal(sb);
			debugLog = sb.ToString();
			return result;
		}

		/// <summary>
		/// 内部共通ロジック。
		/// </summary>
		private string? ScanInternal(StringBuilder log)
		{
			if (_documentRoot == null)
			{
				log?.AppendLine("AttachInternal を実行します。");
				if (!AttachInternal(log))
				{
					log?.AppendLine("AttachInternal 失敗。");
					return null;
				}
			}
			else
			{
				log?.AppendLine("既存の documentRoot を使用します。");
			}

			var doc = _documentRoot ?? _windowRoot;
			if (doc == null)
			{
				log?.AppendLine("documentRoot が null です。");
				return null;
			}

			// --- まず Text 要素全部を調査 -----------------

			AutomationElementCollection texts;
			try
			{
				texts = doc.FindAll(
					TreeScope.Descendants,
					new PropertyCondition(
						AutomationElement.ControlTypeProperty,
						ControlType.Text));
			}
			catch (Exception ex)
			{
				log?.AppendLine("Text 検索中に例外: " + ex.Message);
				return null;
			}

			log?.AppendLine($"Text 要素数: {texts.Count}");

			var candidates = new List<(string text, double bottom)>();

			int logCount = Math.Min(50, texts.Count);
			for (int i = 0; i < texts.Count; i++)
			{
				var el = texts[i];
				string name;
				System.Windows.Rect rect;

				try
				{
					name = el.Current.Name ?? string.Empty;
					rect = el.Current.BoundingRectangle;
				}
				catch
				{
					continue;
				}

				if (i < logCount && log != null)
				{
					var shortName = name;
					if (shortName.Length > 40)
						shortName = shortName.Substring(0, 40) + "…";

					log.AppendLine($"  Text[{i}] Bottom={rect.Bottom}, Name=\"{shortName}\"");
				}

				if (string.IsNullOrWhiteSpace(name))
					continue;

				// 「→」を含むものを候補に
				if (!name.Contains("→"))
					continue;

				// ドル円 っぽいものだけに絞る（ゆるく）
				if (!(name.Contains("ドル") || name.Contains("USD") || name.Contains("ＵＳＤ")))
					continue;

				candidates.Add((name, rect.Bottom));
			}

			log?.AppendLine($"候補メッセージ数: {candidates.Count}");

			if (candidates.Count == 0)
			{
				log?.AppendLine("条件に合う候補がありませんでした。");
				return null;
			}

			var latest = candidates.OrderBy(c => c.bottom).Last();
			log?.AppendLine("最新候補テキスト:");
			log?.AppendLine(latest.text);

			return latest.text;
		}
	}
}
