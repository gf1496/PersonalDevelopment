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

			// POEMS タブにアタッチ
			if (!_poems.Attach())
			{
				BidText.Text = "POEMS 見つからず";
				AskText.Text = "";
				return;
			}

			// Bid / Ask の UI を特定
			_poems.FindBidAsk();

			// 200ms ごとに価格更新
			_timer = new DispatcherTimer();
			_timer.Interval = TimeSpan.FromMilliseconds(200);
			_timer.Tick += Timer_Tick;
			_timer.Start();
		}

		private void Timer_Tick(object? sender, EventArgs e)
		{
			var bid = _poems.GetBid();
			var ask = _poems.GetAsk();

			if (!bid.HasValue || !ask.HasValue)
			{
				// 失敗したら取り直す
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

		/// <summary>
		/// Fincs チャットから最新シグナルを取得し、
		/// 探索ログもダイアログに表示するテスト用ボタン。
		/// </summary>
		private void OnTestReadSignalClick(object sender, RoutedEventArgs e)
		{
			var board = new FincsBoardAutomation();
			string debug;
			var signal = board.ScanDebug(out debug);

			// まずログ全部
			MessageBox.Show(debug, "FINCS DEBUG", MessageBoxButton.OK, MessageBoxImage.Information);

			// ついでに最新シグナルだけ別ダイアログで
			if (string.IsNullOrEmpty(signal))
			{
				MessageBox.Show("シグナルは見つかりませんでした。", "LATEST SIGNAL", MessageBoxButton.OK,
					MessageBoxImage.Information);
			}
			else
			{
				MessageBox.Show(signal, "LATEST SIGNAL", MessageBoxButton.OK, MessageBoxImage.Information);
			}
		}
	}
}
