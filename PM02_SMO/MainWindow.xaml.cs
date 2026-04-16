using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PM02_SMO
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MultiChannel.Checked += Channel_Checked;
            SingleChannel.Checked += Channel_Checked;
        }

        #region Логика СМО (без изменений)
        private void Channel_Checked(object sender, RoutedEventArgs e)
        {
            ChannelInput.Visibility = MultiChannel.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ProblemType_Changed(object sender, SelectionChangedEventArgs e)
        {
            QueueInput.Visibility = Visibility.Collapsed;
            TimeInput.Visibility = Visibility.Collapsed;

            switch (ProblemType.SelectedIndex)
            {
                case 2: QueueInput.Visibility = Visibility.Visible; break;
                case 4: TimeInput.Visibility = Visibility.Visible; break;
            }
        }

        private void ValidateParameters(double lambda, double mu, int n, int problemType)
        {
            if (lambda <= 0 || mu <= 0)
                throw new ArgumentException("Интенсивности должны быть положительными");

            double p = lambda / mu;
            double rho = p / n;

            switch (problemType)
            {
                case 1:
                    if (rho >= 1)
                        throw new ArgumentException($"Система нестабильна: ρ = {rho:F4} >= 1. Увеличьте μ или количество каналов");
                    break;
                case 2:
                    if (!int.TryParse(QueueSize.Text, out int m) || m < 0)
                        throw new ArgumentException("Неверный размер очереди");
                    break;
                case 3:
                    if (p >= n)
                        AddResult("ВНИМАНИЕ:", 0, $"Система может быть нестабильна: p = {p:F4} >= {n}");
                    break;
            }
        }

        private void Calculate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ResultsPanel.Children.Clear();
                ProbabilitiesPanel.Children.Clear();

                double lambda = ParseInput(LambdaInput.Text);
                double mu = 0;
                if (!string.IsNullOrWhiteSpace(MuInput.Text))
                    mu = ParseInput(MuInput.Text);
                else if (!string.IsNullOrWhiteSpace(ServiceTimeInput.Text))
                    mu = 1.0 / ParseInput(ServiceTimeInput.Text);
                else
                    throw new ArgumentException("Задайте либо интенсивность обслуживания (μ), либо время обслуживания (t)");

                int n = 1;
                if (MultiChannel.IsChecked == true)
                {
                    if (!int.TryParse(ChannelCount.Text, out n) || n <= 0)
                        throw new ArgumentException("Неверное количество каналов");
                }

                int maxState = 10;
                if (!int.TryParse(MaxState.Text, out maxState) || maxState <= 0)
                    throw new ArgumentException("Неверное максимальное состояние");

                ValidateParameters(lambda, mu, n, ProblemType.SelectedIndex);

                switch (ProblemType.SelectedIndex)
                {
                    case 0: CalculateWithFailures(lambda, mu, n, maxState); break;
                    case 1: CalculateWithUnlimitedQueue(lambda, mu, n, maxState); break;
                    case 2: CalculateWithLimitedQueue(lambda, mu, n, maxState); break;
                    case 3: CalculateWithFixedTime(lambda, mu, n, maxState); break;
                    case 4: CalculateWithLimitedTime(lambda, mu, n, maxState); break;
                    default: throw new ArgumentException("Выберите тип СМО");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void CalculateWithFailures(double lambda, double mu, int n, int maxState)
        {
            double p = lambda / mu;
            double p0 = 1.0 / (Enumerable.Range(0, n + 1).Sum(k => Math.Pow(p, k) / Factorial(k)));
            double p_otk = Math.Pow(p, n) * p0 / Factorial(n);
            double Q = 1 - p_otk;
            double A = lambda * Q;

            AddResult("Вероятность простоя системы (P0):", p0);
            AddResult("Вероятность отказа (P_otk):", p_otk);
            AddResult("Относительная пропускная способность (Q):", Q);
            AddResult("Абсолютная пропускная способность (A):", A);
            AddResult("Среднее число занятых каналов:", p * Q);

            for (int k = 0; k <= Math.Min(n, maxState); k++)
            {
                double pk = Math.Pow(p, k) * p0 / Factorial(k);
                AddProbability(k, pk);
            }
        }

        private void CalculateWithUnlimitedQueue(double lambda, double mu, int n, int maxState)
        {
            double p = lambda / mu;
            double rho = p / n;
            double p0 = 1.0 / (Enumerable.Range(0, n).Sum(k => Math.Pow(p, k) / Factorial(k)) + Math.Pow(p, n) / (Factorial(n) * (1 - rho)));
            double Lq = Math.Pow(p, n + 1) * p0 / (Factorial(n) * n * Math.Pow(1 - rho, 2));
            double Wq = Lq / lambda;
            double Ws = Wq + 1 / mu;
            double Ls = lambda * Ws;

            AddResult("Вероятность простоя системы (P0):", p0);
            AddResult("Среднее число заявок в очереди (Lq):", Lq);
            AddResult("Среднее время в очереди (Wq):", Wq);
            AddResult("Среднее время в системе (Ws):", Ws);
            AddResult("Среднее число заявок в системе (Ls):", Ls);
            AddResult("Вероятность ожидания (P_ож):", 1 - p0);
            AddResult("Коэффициент загрузки системы (ρ):", rho);

            for (int k = 0; k <= maxState; k++)
            {
                double pk = k <= n ? Math.Pow(p, k) * p0 / Factorial(k) : Math.Pow(p, k) * p0 / (Factorial(n) * Math.Pow(n, k - n));
                AddProbability(k, pk);
            }
        }

        private void CalculateWithLimitedQueue(double lambda, double mu, int n, int maxState)
        {
            if (!int.TryParse(QueueSize.Text, out int m) || m < 0) throw new ArgumentException("Неверный размер очереди");
            double p = lambda / mu;
            double rho = p / n;

            double denominator;
            if (Math.Abs(rho - 1) < 1e-10)
                denominator = Enumerable.Range(0, n).Sum(k => Math.Pow(p, k) / Factorial(k)) + Math.Pow(p, n) / Factorial(n) * (m + 1);
            else
                denominator = Enumerable.Range(0, n).Sum(k => Math.Pow(p, k) / Factorial(k)) + Math.Pow(p, n) / Factorial(n) * (1 - Math.Pow(rho, m + 1)) / (1 - rho);

            double p0 = 1.0 / denominator;
            double p_otk = Math.Pow(p, n + m) * p0 / (Factorial(n) * Math.Pow(n, m));

            double Lq;
            if (Math.Abs(rho - 1) < 1e-10)
                Lq = Math.Pow(p, n) * p0 * m * (m + 1) / (2 * Factorial(n));
            else
                Lq = Math.Pow(p, n) * p0 * rho * (1 - Math.Pow(rho, m) * (m + 1 - m * rho)) / (Factorial(n) * Math.Pow(1 - rho, 2));

            AddResult("Вероятность простоя системы (P0):", p0);
            AddResult("Вероятность отказа (P_otk):", p_otk);
            AddResult("Средняя длина очереди (Lq):", Lq);
            AddResult("Абсолютная пропускная способность (A):", lambda * (1 - p_otk));
            AddResult("Среднее число заявок в системе:", Lq + p * (1 - p_otk));

            for (int k = 0; k <= Math.Min(n + m, maxState); k++)
            {
                double pk = k <= n ? Math.Pow(p, k) * p0 / Factorial(k) : Math.Pow(p, k) * p0 / (Factorial(n) * Math.Pow(n, k - n));
                AddProbability(k, pk);
            }
        }

        private void CalculateWithFixedTime(double lambda, double mu, int n, int maxState)
        {
            double p = lambda / mu;
            double p0 = 1.0 / (Enumerable.Range(0, n + 1).Sum(k => Math.Pow(p, k) / Factorial(k)));
            AddResult("Вероятность простоя системы (P0):", p0);
            AddResult("Примечание:", 0, "Для систем с фиксированным временем требуются специальные модели");
            for (int k = 0; k <= Math.Min(n, maxState); k++) AddProbability(k, Math.Pow(p, k) * p0 / Factorial(k));
        }

        private void CalculateWithLimitedTime(double lambda, double mu, int n, int maxState)
        {
            if (!double.TryParse(MaxTime.Text, out double tau) || tau <= 0) throw new ArgumentException("Неверное время ожидания");
            double p = lambda / mu;
            double p0 = 1.0 / (1 + p * (1 - Math.Exp(-mu * tau)));
            AddResult("Вероятность простоя системы (P0):", p0);
            AddResult("Примечание:", 0, "Упрощенная модель для демонстрации");
            for (int k = 0; k <= maxState; k++) AddProbability(k, k == 0 ? p0 : p0 * Math.Pow(p, k) * (1 - Math.Exp(-mu * tau)));
        }

        private void AddResult(string name, double value)
        {
            ResultsPanel.Children.Add(new TextBlock { Text = $"{name} {value:F6}", Margin = new Thickness(0, 2, 0, 0) });
        }
        private void AddResult(string name, double value, string unit)
        {
            ResultsPanel.Children.Add(new TextBlock { Text = $"{name} {value:F6} {unit}", Margin = new Thickness(0, 2, 0, 0) });
        }
        private void AddProbability(int state, double probability)
        {
            ProbabilitiesPanel.Children.Add(new TextBlock { Text = $"P({state}) = {probability:F6}", Margin = new Thickness(0, 2, 0, 0) });
        }

        private double ParseInput(string text)
        {
            if (double.TryParse(text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result) && result > 0)
                return result;
            throw new ArgumentException("Неверный формат числа или число <= 0");
        }
        private int Factorial(int k) => k == 0 ? 1 : Enumerable.Range(1, k).Aggregate(1, (p, item) => p * item);
        #endregion

        #region Excel Import / Export
        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (ResultsPanel.Children.Count == 0) { MessageBox.Show("Сначала выполните расчет"); return; }

            Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog();
            sfd.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (sfd.ShowDialog() != true) return;

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;

            try
            {
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlWorkbook = xlApp.Workbooks.Add();
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[1];

                int row = 1;
                xlWorksheet.Cells[row, 1] = "ПАРАМЕТРЫ ВВОДА"; xlWorksheet.Cells[row, 2] = "ЗНАЧЕНИЕ"; row++;

                xlWorksheet.Cells[row, 1] = "Тип СМО"; xlWorksheet.Cells[row, 2] = ((ComboBoxItem)ProblemType.SelectedItem).Content.ToString(); row++;
                xlWorksheet.Cells[row, 1] = "Тип каналов"; xlWorksheet.Cells[row, 2] = MultiChannel.IsChecked == true ? "Многоканальная" : "Одноканальная"; row++;
                xlWorksheet.Cells[row, 1] = "Интенсивность потока (λ)"; xlWorksheet.Cells[row, 2] = LambdaInput.Text; row++;
                xlWorksheet.Cells[row, 1] = "Интенсивность обслуживания (μ)"; xlWorksheet.Cells[row, 2] = MuInput.Text; row++;
                xlWorksheet.Cells[row, 1] = "Время обслуживания (t)"; xlWorksheet.Cells[row, 2] = ServiceTimeInput.Text; row++;
                xlWorksheet.Cells[row, 1] = "Количество каналов (n)"; xlWorksheet.Cells[row, 2] = ChannelCount.Text; row++;
                xlWorksheet.Cells[row, 1] = "Размер очереди (m)"; xlWorksheet.Cells[row, 2] = QueueSize.Text; row++;
                xlWorksheet.Cells[row, 1] = "Время ожидания (τ)"; xlWorksheet.Cells[row, 2] = MaxTime.Text; row++;
                xlWorksheet.Cells[row, 1] = "Макс. состояние"; xlWorksheet.Cells[row, 2] = MaxState.Text; row++;

                row++; // Пустая строка
                xlWorksheet.Cells[row, 1] = "ОСНОВНЫЕ ПОКАЗАТЕЛИ"; row++;
                xlWorksheet.Cells[row, 1] = "Показатель"; xlWorksheet.Cells[row, 2] = "Значение"; row++;

                foreach (var child in ResultsPanel.Children)
                {
                    if (child is TextBlock tb)
                    {
                        string text = tb.Text;
                        int lastColon = text.LastIndexOf(':');
                        if (lastColon > 0)
                        {
                            xlWorksheet.Cells[row, 1] = text.Substring(0, lastColon).Trim();
                            xlWorksheet.Cells[row, 2] = text.Substring(lastColon + 1).Trim();
                        }
                        else
                        {
                            xlWorksheet.Cells[row, 1] = text;
                        }
                        row++;
                    }
                }

                row++; // Пустая строка
                xlWorksheet.Cells[row, 1] = "ВЕРОЯТНОСТИ СОСТОЯНИЙ"; row++;
                xlWorksheet.Cells[row, 1] = "Состояние"; xlWorksheet.Cells[row, 2] = "Вероятность"; row++;

                foreach (var child in ProbabilitiesPanel.Children)
                {
                    if (child is TextBlock tb)
                    {
                        string text = tb.Text.Replace("P(", "").Replace(")", "");
                        string[] parts = text.Split('=');
                        if (parts.Length == 2)
                        {
                            xlWorksheet.Cells[row, 1] = parts[0].Trim();
                            xlWorksheet.Cells[row, 2] = parts[1].Trim();
                        }
                        row++;
                    }
                }

                // Автоподбор ширины
                xlWorksheet.Columns.AutoFit();

                xlWorkbook.SaveAs(sfd.FileName);
                MessageBox.Show("Успешно экспортировано!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Ошибка экспорта: {ex.Message}"); }
            finally { ReleaseExcelObject(null, xlWorksheet, xlWorkbook, xlApp); xlWorksheet = null; xlWorkbook = null; xlApp = null; }
        }

        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            if (ofd.ShowDialog() != true) return;

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    string param = xlRange.Cells[i, 1].Value?.ToString().Trim();
                    string val = xlRange.Cells[i, 2].Value?.ToString().Trim();

                    if (string.IsNullOrWhiteSpace(param)) continue;

                    switch (param)
                    {
                        case "Тип СМО":
                            for (int j = 0; j < ProblemType.Items.Count; j++)
                                if (((ComboBoxItem)ProblemType.Items[j]).Content.ToString() == val) { ProblemType.SelectedIndex = j; break; }
                            break;
                        case "Тип каналов":
                            if (val == "Многоканальная") MultiChannel.IsChecked = true;
                            else SingleChannel.IsChecked = true;
                            break;
                        case "Интенсивность потока (λ)": LambdaInput.Text = val ?? ""; break;
                        case "Интенсивность обслуживания (μ)": MuInput.Text = val ?? ""; break;
                        case "Время обслуживания (t)": ServiceTimeInput.Text = val ?? ""; break;
                        case "Количество каналов (n)": ChannelCount.Text = val ?? "2"; break;
                        case "Размер очереди (m)": QueueSize.Text = val ?? "5"; break;
                        case "Время ожидания (τ)": MaxTime.Text = val ?? ""; break;
                        case "Макс. состояние": MaxState.Text = val ?? "10"; break;
                    }

                    // Как только дошли до результатов - прекращаем чтение параметров
                    if (param == "ОСНОВНЫЕ ПОКАЗАТЕЛИ") break;
                }

                MessageBox.Show("Параметры успешно загружены.\nНажмите 'Рассчитать' для получения результатов.", "Импорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Ошибка импорта:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }
            finally { ReleaseExcelObject(xlRange, xlWorksheet, xlWorkbook, xlApp); xlRange = null; xlWorksheet = null; xlWorkbook = null; xlApp = null; }
        }

        private void ReleaseExcelObject(object xlRange = null, object xlWorksheet = null, object xlWorkbook = null, object xlApp = null)
        {
            try
            {
                if (xlRange != null) { Marshal.ReleaseComObject(xlRange); xlRange = null; }
                if (xlWorksheet != null) { Marshal.ReleaseComObject(xlWorksheet); xlWorksheet = null; }
                if (xlWorkbook != null) { ((Excel.Workbook)xlWorkbook).Close(false); Marshal.ReleaseComObject(xlWorkbook); xlWorkbook = null; }
                if (xlApp != null) { ((Excel.Application)xlApp).Quit(); Marshal.ReleaseComObject(xlApp); xlApp = null; }
            }
            catch { }
            finally { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers(); }
        }
        #endregion
    }
}