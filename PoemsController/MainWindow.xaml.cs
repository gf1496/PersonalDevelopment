using System;
using System.Windows;
using System.Windows.Threading;

namespace PoemsController
{
	public partial class MainWindow : Window
	{
		private PoemsAutomation _poems;
		private DispatcherTimer _timer;

		public MainWindow()
		{
			InitializeComponent();

			_poems = new PoemsAutomation();

			// 1. Chrome の POEMS ウィンドウにアタッチ
			if (!_poems.Attach())
			{
				BidText.Text = "POEMS 見つからず";
				AskText.Text = "";
				return;
			}

			// 2. 最初の Bid / Ask を特定
			_poems.FindBidAsk();

			// 3. 200ms ごとに値を更新
			_timer = new DispatcherTimer();
			_timer.Interval = TimeSpan.FromMilliseconds(200);
			_timer.Tick += Timer_Tick;
			_timer.Start();
		}

		private void Timer_Tick(object sender, EventArgs e)
		{
			var bid = _poems.GetBid();
			var ask = _poems.GetAsk();

			// 値が取れなくなっていたら（要素が再生成されている可能性）
			if (!bid.HasValue || !ask.HasValue)
			{
				// Bid/Ask の要素を再検索
				_poems.FindBidAsk();

				// もう一度読み直す
				bid = _poems.GetBid();
				ask = _poems.GetAsk();
			}

			if (bid.HasValue)
				BidText.Text = bid.Value.ToString("F3");
			else
				BidText.Text = "(no bid)";

			if (ask.HasValue)
				AskText.Text = ask.Value.ToString("F3");
			else
				AskText.Text = "(no ask)";
		}

		private void OnBuyClick(object sender, RoutedEventArgs e)
		{
			if (double.TryParse(LotBox.Text, out var lot))
			{
				_poems.SetLot(lot);
				_poems.ClickBuy();
			}
		}

		private void OnSellClick(object sender, RoutedEventArgs e)
		{
			if (double.TryParse(LotBox.Text, out var lot))
			{
				_poems.SetLot(lot);
				_poems.ClickSell();
			}
		}
	}
}
