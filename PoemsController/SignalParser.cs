using System.Text;
using System.Text.RegularExpressions;

namespace PoemsController
{
	public enum TradeAction
	{
		None,
		LongEntry,
		ShortEntry,
		Close
	}

	public enum CloseReason
	{
		None,
		TakeProfit,    // 利確
		StopLoss,      // 損切り（全決済）
		PartialRecent  // 直近◯割分だけ損切り
	}

	public class ParsedSignal
	{
		public TradeAction Action { get; set; } = TradeAction.None;

		/// <summary>
		/// エントリー時: 「最大ロットの◯割」の ◯
		/// 直近部分損切り時: 「直近◯割分だけ損切り」の ◯
		/// </summary>
		public int? Ratio { get; set; }

		public bool IsAdd { get; set; }

		public CloseReason CloseReason { get; set; } = CloseReason.None;

		public override string ToString()
		{
			return $"Action={Action}, Ratio={Ratio}, IsAdd={IsAdd}, CloseReason={CloseReason}";
		}
	}

	public static class SignalParser
	{
		// 全角数字・括弧を半角に寄せる
		private static string Normalize(string text)
		{
			if (text == null) return string.Empty;

			var sb = new StringBuilder(text.Length);
			foreach (var ch in text)
			{
				if (ch >= '０' && ch <= '９')
				{
					sb.Append((char)('0' + (ch - '０')));
				}
				else if (ch == '（') sb.Append('(');
				else if (ch == '）') sb.Append(')');
				else sb.Append(ch);
			}
			return sb.ToString();
		}

		// エントリー: 最大ロットの◯割でロング/ショートエントリー
		private static readonly Regex EntryRegex =
			new Regex(@"最大ロットの\s*([0-9]+)\s*割で\s*(ロング|ショート)エントリー",
					  RegexOptions.Compiled);

		// 直近部分損切り: 直近◯割分だけ損切り
		private static readonly Regex PartialCloseRegex =
			new Regex(@"直近\s*([0-9]+)\s*割分だけ損切り",
					  RegexOptions.Compiled);

		public static ParsedSignal ParseSignal(string rawText)
		{
			var result = new ParsedSignal();

			if (string.IsNullOrWhiteSpace(rawText))
				return result;

			var text = Normalize(rawText);

			// 「→」より右側だけを見る
			var arrowIndex = text.IndexOf('→');
			if (arrowIndex >= 0 && arrowIndex + 1 < text.Length)
			{
				text = text[(arrowIndex + 1)..];
			}

			// 1) 直近◯割分だけ損切り
			var partialMatch = PartialCloseRegex.Match(text);
			if (partialMatch.Success)
			{
				result.Action = TradeAction.Close;
				result.CloseReason = CloseReason.PartialRecent;

				if (int.TryParse(partialMatch.Groups[1].Value, out var r))
					result.Ratio = r;

				return result;
			}

			// 2) 利確
			if (text.Contains("利確"))
			{
				result.Action = TradeAction.Close;
				result.CloseReason = CloseReason.TakeProfit;
				return result;
			}

			// 3) 通常損切り
			if (text.Contains("損切り"))
			{
				result.Action = TradeAction.Close;
				result.CloseReason = CloseReason.StopLoss;
				return result;
			}

			// 4) エントリー
			var m = EntryRegex.Match(text);
			if (m.Success)
			{
				if (int.TryParse(m.Groups[1].Value, out var ratio))
					result.Ratio = ratio;

				var dir = m.Groups[2].Value;
				if (dir.Contains("ロング"))
					result.Action = TradeAction.LongEntry;
				else if (dir.Contains("ショート"))
					result.Action = TradeAction.ShortEntry;

				if (text.Contains("追加"))
					result.IsAdd = true;

				return result;
			}

			// 5) 何にもマッチしない
			return result;
		}
	}
}
