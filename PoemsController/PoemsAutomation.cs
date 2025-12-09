using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;

namespace PoemsController
{
	public class PoemsAutomation
	{
		private AutomationElement _root;

		// 価格表示
		private AutomationElement _bidElement;
		private AutomationElement _askElement;

		// 注文系
		private AutomationElement _lotElement;
		private AutomationElement _buyElement;
		private AutomationElement _sellElement;

		// ========= Win32 マウス操作 =========

		[DllImport("user32.dll")]
		private static extern bool SetCursorPos(int X, int Y);

		[DllImport("user32.dll")]
		private static extern void mouse_event(
			uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

		private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
		private const uint MOUSEEVENTF_LEFTUP = 0x0004;

		private void ClickElementCenter(AutomationElement elem)
		{
			if (elem == null) return;

			System.Windows.Rect r;
			try { r = elem.Current.BoundingRectangle; }
			catch { return; }

			if (r.IsEmpty) return;

			int x = (int)(r.Left + r.Width / 2.0);
			int y = (int)(r.Top + r.Height / 2.0);

			SetCursorPos(x, y);
			mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
			mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
		}

		// =======================
		// Chrome / POEMS へのアタッチ
		// =======================

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

		// =======================
		// Bid / Ask の特定 & 取得
		// =======================

		public void FindBidAsk()
		{
			if (_root == null) return;

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

				if (!double.TryParse(
						name,
						NumberStyles.Float,
						CultureInfo.InvariantCulture,
						out var value))
				{
					continue;
				}

				if (value < 50.0 || value > 200.0)
					continue;

				if (HasLabelAncestor(t, bidKeywords))
					bidCandidates.Add(t);

				if (HasLabelAncestor(t, askKeywords))
					askCandidates.Add(t);
			}

			_bidElement = bidCandidates.Count > 0 ? bidCandidates[0] : null;
			_askElement = askCandidates.Count > 0 ? askCandidates[0] : null;
		}

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

		// =======================
		// 数量入力欄の特定
		// =======================

		private AutomationElement FindQuantityEdit()
		{
			if (_root == null) return null;

			// 「数量」ラベルを Text から探す
			var textCond = new PropertyCondition(
				AutomationElement.ControlTypeProperty, ControlType.Text);

			var texts = _root.FindAll(TreeScope.Descendants, textCond)
							 .Cast<AutomationElement>()
							 .ToList();

			AutomationElement qtyLabel = null;

			foreach (var t in texts)
			{
				string name;
				try { name = t.Current.Name ?? string.Empty; }
				catch { continue; }

				if (string.IsNullOrEmpty(name))
					continue;

				if (name.Contains("数量") ||
					name.IndexOf("Quantity", StringComparison.OrdinalIgnoreCase) >= 0 ||
					name.IndexOf("Qty", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					qtyLabel = t;
					break;
				}
			}

			if (qtyLabel == null)
				return null;

			var labelRect = qtyLabel.Current.BoundingRectangle;
			double anchorY = labelRect.Top;

			var editCond = new PropertyCondition(
				AutomationElement.ControlTypeProperty, ControlType.Edit);

			var edits = _root.FindAll(TreeScope.Descendants, editCond)
							 .Cast<AutomationElement>()
							 .ToList();

			AutomationElement best = null;
			double bestScore = double.MaxValue;

			foreach (var e in edits)
			{
				System.Windows.Rect r;
				try { r = e.Current.BoundingRectangle; }
				catch { continue; }

				if (r.IsEmpty) continue;

				double dy = Math.Abs(r.Top - anchorY);
				if (dy > 120.0) continue; // 数量行付近だけ

				double dx = Math.Max(0.0, labelRect.Left - r.Left);
				double score = dy + dx * 0.1;

				if (score < bestScore)
				{
					bestScore = score;
					best = e;
				}
			}

			return best;
		}

		// =======================
		// 買・売要素の特定（数量欄の近所だけ見る）
		// =======================

		private AutomationElement FindNeighborByName(string keyword, AutomationElement anchor)
		{
			if (_root == null || anchor == null) return null;

			var refRect = anchor.Current.BoundingRectangle;

			var typeCond = new OrCondition(
				new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
				new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
				new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink),
				new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
			);

			var elems = _root.FindAll(TreeScope.Descendants, typeCond)
							 .Cast<AutomationElement>()
							 .ToList();

			AutomationElement best = null;
			double bestScore = double.MaxValue;

			foreach (var e in elems)
			{
				string name;
				System.Windows.Rect r;

				try
				{
					name = e.Current.Name ?? string.Empty;
					r = e.Current.BoundingRectangle;
				}
				catch
				{
					continue;
				}

				if (string.IsNullOrEmpty(name)) continue;
				if (!name.Contains(keyword) &&
					name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0)
					continue;

				if (r.IsEmpty) continue;

				// 数量欄の縦位置からあまり離れていない要素だけ見る
				double dy = Math.Abs(r.Top - refRect.Top);
				if (dy > 150.0) continue;

				// 数量欄より右側にあるものを優先
				double dx = Math.Max(0.0, r.Left - refRect.Left);

				double score = dy * 2.0 + dx; // まず縦が近い、次に右側のもの

				if (score < bestScore)
				{
					bestScore = score;
					best = e;
				}
			}

			return best;
		}

		public void FindOrderControls()
		{
			_lotElement = FindQuantityEdit();
			if (_lotElement == null)
			{
				MessageBox.Show("数量入力欄(Edit)が見つかりませんでした。", "FindOrderControls");
				return;
			}

			_buyElement = FindNeighborByName("買", _lotElement);
			_sellElement = FindNeighborByName("売", _lotElement);

			if (_buyElement == null)
				MessageBox.Show("買要素が見つかりませんでした。", "FindOrderControls");

			if (_sellElement == null)
				MessageBox.Show("売要素が見つかりませんでした。", "FindOrderControls");
		}

		// =======================
		// 操作系 API
		// =======================

		public void SetLot(int lot)
		{
			if (_lotElement == null)
				FindOrderControls();

			if (_lotElement == null)
			{
				MessageBox.Show("ロット入力欄が特定できていません。", "SetLot");
				return;
			}

			if (_lotElement.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern))
			{
				((ValuePattern)pattern).SetValue(lot.ToString(CultureInfo.InvariantCulture));
			}
			else
			{
				MessageBox.Show("数量欄が ValuePattern をサポートしていません。", "SetLot");
			}
		}

		public void ClickBuy()
		{
			if (_buyElement == null)
				FindOrderControls();

			if (_buyElement == null)
			{
				MessageBox.Show("買要素が特定できていません。", "ClickBuy");
				return;
			}

			if (_buyElement.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
			{
				((InvokePattern)pattern).Invoke();
			}
			else
			{
				// InvokePattern が無い → 中心をクリックする
				ClickElementCenter(_buyElement);
			}
		}

		public void ClickSell()
		{
			if (_sellElement == null)
				FindOrderControls();

			if (_sellElement == null)
			{
				MessageBox.Show("売要素が特定できていません。", "ClickSell");
				return;
			}

			if (_sellElement.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
			{
				((InvokePattern)pattern).Invoke();
			}
			else
			{
				// InvokePattern が無い → 中心をクリックする
				ClickElementCenter(_sellElement);
			}
		}
	}
}
