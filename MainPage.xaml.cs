using System.Data;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Maui.Controls;
using Grid = Microsoft.Maui.Controls.Grid;

namespace MyExcelMAUIApp
{
    public partial class MainPage : ContentPage
    {
        const int CountColumn = 20;
        const int CountRow = 50;

        public MainPage()
        {
            InitializeComponent();
            Loaded += OnPageLoaded;
        }

        private void OnPageLoaded(object sender, EventArgs e)
        {
            CreateGrid();
        }

        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        private void AddColumnsAndColumnLabels()
        {
            for (int col = 0; col < CountColumn + 1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                if (col > 0)
                {
                    var label = new Label
                    {
                        Text = GetColumnName(col),
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0);
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                }
            }
        }

        private void AddRowsAndCellEntries()
        {
            for (int row = 0; row < CountRow; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());

                var label = new Label
                {
                    Text = (row + 1).ToString(),
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                grid.Children.Add(label);

                for (int col = 0; col < CountColumn; col++)
                {
                    var entry = new Entry
                    {
                        Text = "",
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    entry.Unfocused += Entry_Unfocused;
                    Grid.SetRow(entry, row + 1);
                    Grid.SetColumn(entry, col + 1);
                    grid.Children.Add(entry);
                }
            }
        }

        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private void Entry_Unfocused(object sender, FocusEventArgs e)
        {
            if (sender is Entry entry && !string.IsNullOrWhiteSpace(entry.Text))
            {
                string text = entry.Text.Trim();
                if (text.StartsWith("="))
                {
                    double result = EvaluateExpression(text);
                    if (!double.IsNaN(result))
                        entry.Text = result.ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        private void CalculateButton_Clicked(object sender, EventArgs e)
        {
            int errors = 0;

            foreach (var child in grid.Children)
            {
                if (child is Entry entry && !string.IsNullOrWhiteSpace(entry.Text))
                {
                    string text = entry.Text.Trim();
                    if (text.StartsWith("="))
                    {
                        double result = EvaluateExpression(text);
                        if (double.IsNaN(result))
                        {
                            entry.TextColor = Colors.Red;
                            errors++;
                        }
                        else
                        {
                            entry.TextColor = Colors.Black;
                            entry.Text = result.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                }
            }

            if (errors == 0)
                DisplayAlert("Результат", "Усі вирази успішно обчислено!", "OK");
            else
                DisplayAlert("Помилка", $"Виявлено {errors} помилок у формулах!", "OK");
        }


        private void SaveButton_Clicked(object sender, EventArgs e)
        {
            var lines = new List<string>();
            for (int row = 0; row < CountRow; row++)
            {
                var rowValues = new List<string>();
                for (int col = 0; col < CountColumn; col++)
                {
                    var entry = GetEntryAt(row, col);
                    rowValues.Add(entry?.Text ?? "");
                }
                lines.Add(string.Join(";", rowValues));
            }
            File.WriteAllLines("table.csv", lines);
            DisplayAlert("Збережено", "Дані збережено у файл table.csv", "OK");
        }

        private void ReadButton_Clicked(object sender, EventArgs e)
        {
            if (!File.Exists("table.csv"))
            {
                DisplayAlert("Помилка", "Файл не знайдено!", "OK");
                return;
            }
            var lines = File.ReadAllLines("table.csv");
            for (int row = 0; row < Math.Min(lines.Length, CountRow); row++)
            {
                var values = lines[row].Split(';');
                for (int col = 0; col < Math.Min(values.Length, CountColumn); col++)
                {
                    var entry = GetEntryAt(row, col);
                    if (entry != null)
                        entry.Text = values[col];
                }
            }
            DisplayAlert("Зчитано", "Дані зчитано з table.csv", "OK");
        }

        private Entry? GetEntryAt(int row, int col)
        {
            foreach (var child in grid.Children)
            {
                if (child is Entry entry &&
                    Grid.GetRow(entry) == row + 1 &&
                    Grid.GetColumn(entry) == col + 1)
                    return entry;
            }
            return null;
        }

        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                System.Environment.Exit(0);
            }
        }

        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студентки Марії Балєєвої", "OK");
        }

        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(lastRowIndex);

                var toRemove = new List<Microsoft.Maui.Controls.View>();

                foreach (var child in grid.Children)
                {
                    if (child is Microsoft.Maui.Controls.View view &&
                        Grid.GetRow(view) == lastRowIndex)
                    {
                        toRemove.Add(view);
                    }
                }

                foreach (var view in toRemove)
                {
                    grid.Children.Remove(view);
                }
            }
        }

        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;
                grid.ColumnDefinitions.RemoveAt(lastColumnIndex);

                var toRemove = new List<Microsoft.Maui.Controls.View>();

                foreach (var child in grid.Children)
                {
                    if (child is Microsoft.Maui.Controls.View view &&
                        Grid.GetColumn(view) == lastColumnIndex)
                    {
                        toRemove.Add(view);
                    }
                }

                foreach (var view in toRemove)
                {
                    grid.Children.Remove(view);
                }
            }
        }


        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            int newRow = grid.RowDefinitions.Count;
            grid.RowDefinitions.Add(new RowDefinition());

            var label = new Label
            {
                Text = newRow.ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, newRow);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);

            for (int col = 0; col < CountColumn; col++)
            {
                var entry = new Entry();
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                grid.Children.Add(entry);
            }
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count;
            grid.ColumnDefinitions.Add(new ColumnDefinition());

            var label = new Label
            {
                Text = GetColumnName(newColumn),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, newColumn);
            grid.Children.Add(label);

            for (int row = 0; row < CountRow; row++)
            {
                var entry = new Entry();
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                grid.Children.Add(entry);
            }
        }
       

        private double EvaluateExpression(string expression, HashSet<string>? visited = null)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return double.NaN;

            visited ??= new HashSet<string>();

         
            expression = expression.Trim();
            if (expression.StartsWith("="))
                expression = expression.Substring(1);
            expression = expression.Replace(" ", "");

            try
            {
        
                string replaced = ReplaceCellReferences(expression, visited);

               
                replaced = EvaluateFunctions(replaced, visited);
    
       
                var dt = new System.Data.DataTable();
                var raw = dt.Compute(replaced, null);
                double result = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return result;


            }
            catch
            {
                return double.NaN;
            }
        }

        
        private string ReplaceCellReferences(string expr, HashSet<string> visited)
        {
            var regex = new Regex(@"([A-Za-z]+)(\d+)");
            string evaluator(Match m)
            {
                string colLetters = m.Groups[1].Value.ToUpperInvariant();
                string rowNumberStr = m.Groups[2].Value;
                if (!int.TryParse(rowNumberStr, out int rowNumber))
                    return "0";

                int colIndex = ColumnNameToIndex(colLetters);
                int rowIndex = rowNumber - 1;
                if (rowIndex < 0 || colIndex < 0)
                    return "0";

                string key = $"{rowIndex}:{colIndex}";
                if (visited.Contains(key))
                    return "0";

                var entry = GetEntryAt(rowIndex, colIndex);
                string val = entry?.Text ?? "0";
                if (string.IsNullOrWhiteSpace(val))
                    return "0";

                if (val.TrimStart().StartsWith("="))
                {
                    visited.Add(key);
                    double sub = EvaluateExpression(val, visited);
                    visited.Remove(key);
                    if (double.IsNaN(sub))
                        return "0";
                    return sub.ToString(CultureInfo.InvariantCulture);
                }

             
                if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                    return parsed.ToString(CultureInfo.InvariantCulture);

                
                return "0";
            }

            string replaced = regex.Replace(expr, new MatchEvaluator(m => evaluator(m)));
            return replaced;
        }

        // Обробка спеціальних функцій: inc, dec, mmax, mmin, not
        

        private string EvaluateFunctions(string expr, HashSet<string> visited)
        {
           
            var funcRegex = new Regex(@"\b(inc|dec|mmax|mmin|not)\s*\(([^()]*?)\)", RegexOptions.IgnoreCase);

         
            while (funcRegex.IsMatch(expr))
            {
                expr = funcRegex.Replace(expr, new MatchEvaluator(m =>
                {
                    string fname = m.Groups[1].Value.ToLowerInvariant();
                    string argsRaw = m.Groups[2].Value.Trim();

                  
                    string[] args = string.IsNullOrWhiteSpace(argsRaw)
                        ? new string[0]
                        : argsRaw.Split(',').Select(s => s.Trim()).ToArray();

                    try
                    {
                        switch (fname)
                        {
                            case "inc":
                                {
                                    if (args.Length != 1) return "0";
                                    double v = EvaluateExpression(args[0], visited);
                                    if (double.IsNaN(v)) return "0";
                                    return (v + 1).ToString(CultureInfo.InvariantCulture);
                                }
                            case "dec":
                                {
                                    if (args.Length != 1) return "0";
                                    double v = EvaluateExpression(args[0], visited);
                                    if (double.IsNaN(v)) return "0";
                                    return (v - 1).ToString(CultureInfo.InvariantCulture);
                                }
                            case "mmax":
                                {
                                    if (args.Length < 1) return "0";
                                    double[] vals = args.Select(a => EvaluateExpression(a, visited)).ToArray();
                                    if (vals.Any(d => double.IsNaN(d))) return "0";
                                    double mx = vals.Max();
                                    return mx.ToString(CultureInfo.InvariantCulture);
                                }
                            case "mmin":
                                {
                                    if (args.Length < 1) return "0";
                                    double[] vals = args.Select(a => EvaluateExpression(a, visited)).ToArray();
                                    if (vals.Any(d => double.IsNaN(d))) return "0";
                                    double mn = vals.Min();
                                    return mn.ToString(CultureInfo.InvariantCulture);
                                }
                            case "not":
                                {
                                    if (args.Length != 1) return "0";
                                    double v = EvaluateExpression(args[0], visited);
                                    if (double.IsNaN(v)) return "0";
                               
                                    return ((Math.Abs(v) < 1e-12) ? 1.0 : 0.0).ToString(CultureInfo.InvariantCulture);
                                }
                            default:
                                return "0";
                        }
                    }
                    catch
                    {
                        return "0";
                    }
                }), 1); 
            }

            return expr;
        }

       
        private int ColumnNameToIndex(string name)
        {
            if (string.IsNullOrEmpty(name)) return -1;
            int result = 0;
            foreach (char c in name)
            {
                if (c < 'A' || c > 'Z') return -1;
                result = result * 26 + (c - 'A' + 1);
            }
            return result - 1;
        }

    }
}
