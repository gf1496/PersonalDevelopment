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

			if (!_poems.Attach())
			{
				BidText.Text = "POEMS 見つからず";
				AskText.Text = "";
				return;
			}

			_poems.FindBidAsk();

			_timer = new DispatcherTimer();
			_timer.Interval = TimeSpan.FromMilliseconds(200);
			_timer.Tick += Timer_Tick;
			_timer.Start();
		}

		private void Timer_Tick(object sender, EventArgs e)
		{
			var bid = _poems.GetBid();
			var ask = _poems.GetAsk();

			if (!bid.HasValue || !ask.HasValue)
			{
				_poems.FindBidAsk();
				bid = _poems.GetBid();
				ask = _poems.GetAsk();
			}

			BidText.Text = bid.HasValue ? bid.Value.ToString("F3") : "(no bid)";
			AskText.Text = ask.HasValue ? ask.Value.ToString("F3") : "(no ask)";
		}

		private void OnBuyClick(object sender, RoutedEventArgs e)
		{
			if (int.TryParse(LotBox.Text, out var lot))
			{
				_poems.SetLot(lot);
				_poems.ClickBuy();
			}
		}

		private void OnSellClick(object sender, RoutedEventArgs e)
		{
			if (int.TryParse(LotBox.Text, out var lot))
			{
				_poems.SetLot(lot);
				_poems.ClickSell();
			}
		}
	}
}
