using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Automation;

namespace PoemsController
{
	public class PoemsAutomation
	{
		private AutomationElement _root;

		private AutomationElement _bidElement;
		private AutomationElement _askElement;

		private AutomationElement _lotTextBox;
		private AutomationElement _buyButton;
		private AutomationElement _sellButton;

		/// <summary>
		/// Chrome 上の POEMS ウィンドウにアタッチする
		///   - ウィンドウタイトルに "POEMS" を含む
		///   - かつプロセス名が chrome のものだけを対象
		/// </summary>
		public bool Attach()
		{
			var desktop = AutomationElement.RootElement;
			if (desktop == null)
				return false;

			var windowCondition = new PropertyCondition(
				AutomationElement.ControlTypeProperty, ControlType.Window);

			var windows = desktop.FindAll(TreeScope.Children, windowCondition);

			_root = null;

			foreach (AutomationElement win in windows)
			{
				string title;
				int pid;

				try
				{
					title = win.Current.Name ?? string.Empty;
					pid = win.Current.ProcessId;
				}
				catch
				{
					continue;
				}

				if (title.IndexOf("POEMS", StringComparison.OrdinalIgnoreCase) < 0)
					continue;

				string procName;
				try
				{
					var proc = Process.GetProcessById(pid);
					procName = proc.ProcessName;
				}
				catch
				{
					continue;
				}

				if (!procName.Contains("chrome", StringComparison.OrdinalIgnoreCase))
					continue;

				_root = win;
				break;
			}

			return _root != null;
		}

		/// <summary>
		/// Bid / Ask 要素を特定する。
		/// 戦略:
		///   1) Document 内の Text を全列挙
		///   2) 数値(50〜200)に変換できるものを候補に
		///   3) 各候補について親方向にさかのぼり、
		///      「ビッド / Bid」「オファー / Offer」を含む祖先がいるか判定
		///   4) ビッド系の祖先を持つもの → Bid 候補
		///      オファー系の祖先を持つもの → Ask 候補
		/// </summary>
		public void FindBidAsk()
		{
			if (_root == null) return;

			// "POEMS 2.0" ドキュメントを起点にする
			var docCond = new PropertyCondition(
				AutomationElement.ControlTypeProperty, ControlType.Document);

			var doc = _root.FindFirst(TreeScope.Descendants, docCond) ?? _root;

			var textCond = new PropertyCondition(
				AutomationElement.ControlTypeProperty, ControlType.Text);

			var texts = doc.FindAll(TreeScope.Descendants, textCond);

			var bidKeywords = new[] { "ビッド", "Bid" };
			var askKeywords = new[] { "オファー", "Offer" };

			var bidCandidates = new List<AutomationElement>();
			var askCandidates = new List<AutomationElement>();

			foreach (AutomationElement t in texts)
			{
				string name;
				try
				{
					name = t.Current.Name ?? string.Empty;
				}
				catch
				{
					continue;
				}

				// 数字でないものは対象外
				if (!double.TryParse(
						name,
						NumberStyles.Float,
						CultureInfo.InvariantCulture,
						out var value))
				{
					continue;
				}

				// USD/JPY レートっぽい値だけ
				if (value < 50.0 || value > 200.0)
					continue;

				if (HasLabelAncestor(t, bidKeywords))
					bidCandidates.Add(t);

				if (HasLabelAncestor(t, askKeywords))
					askCandidates.Add(t);
			}

			// 一旦、最初に見つかった候補を採用（通常は1件ずつのはず）
			_bidElement = bidCandidates.Count > 0 ? bidCandidates[0] : null;
			_askElement = askCandidates.Count > 0 ? askCandidates[0] : null;
		}

		/// <summary>
		/// elem の親方向にさかのぼっていき、Name に指定キーワードのどれかを含む要素があるか。
		/// </summary>
		private bool HasLabelAncestor(AutomationElement elem, string[] keywords)
		{
			if (elem == null) return false;

			var walker = TreeWalker.ControlViewWalker;
			var current = walker.GetParent(elem);

			while (current != null)
			{
				string name;
				try
				{
					name = current.Current.Name ?? string.Empty;
				}
				catch
				{
					name = string.Empty;
				}

				foreach (var kw in keywords)
				{
					if (!string.IsNullOrEmpty(kw) &&
						name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
					{
						return true;
					}
				}

				current = walker.GetParent(current);
			}

			return false;
		}

		public double? GetBid()
		{
			if (_bidElement == null) return null;

			try
			{
				var s = _bidElement.Current.Name;
				if (double.TryParse(
						s,
						NumberStyles.Float,
						CultureInfo.InvariantCulture,
						out var v))
				{
					return v;
				}
			}
			catch { }

			return null;
		}

		public double? GetAsk()
		{
			if (_askElement == null) return null;

			try
			{
				var s = _askElement.Current.Name;
				if (double.TryParse(
						s,
						NumberStyles.Float,
						CultureInfo.InvariantCulture,
						out var v))
				{
					return v;
				}
			}
			catch { }

			return null;
		}

		// 以下はまだダミー（あとでロット入力＆ボタンクリックを実装）

		public void SetLot(double lot)
		{
			if (_lotTextBox == null) return;

			if (_lotTextBox.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern))
			{
				((ValuePattern)pattern).SetValue(lot.ToString(CultureInfo.InvariantCulture));
			}
		}

		public void ClickBuy()
		{
			if (_buyButton == null) return;

			if (_buyButton.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
			{
				((InvokePattern)pattern).Invoke();
			}
		}

		public void ClickSell()
		{
			if (_sellButton == null) return;

			if (_sellButton.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
			{
				((InvokePattern)pattern).Invoke();
			}
		}
	}
}
