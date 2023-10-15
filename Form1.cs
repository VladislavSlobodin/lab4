using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDX.Direct2D1;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace lab4;

public partial class Form1 : Form
{
    #region WordVariables
    private List<string> _wordLabels = new()
    {
        "Имя_Отчество",
        "Событие",
        "Дата",
        "Адрес",
        "Имя_Отчество_2",
    };
    private Dictionary<string, string> _keyValuePairs = new();
    private string _examplePath = string.Empty;
    private static readonly string _filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*";
    #endregion
    #region ExcelVariables
    private List<string> _excelLabels = new()
    {
        "Регион",
        "Лыжи",
        "Коньки",
        "Сани",
        "Всего",
    };
    private string _excelDocumentPath = string.Empty;
    #endregion

    public Form1()
    {
        InitializeComponent();
        InitializeWordTable();
        InitializeExcelTable();
    }

    #region Word
    private void InitializeWordTable()
    {
        WordDataGridView.Columns.AddRange(new DataGridViewTextBoxColumn[] { new() { HeaderText = "Шаблон", Name = "ExampleColumn" }, new() { HeaderText = "Значение", Name = "ValueColumn" } });
        _wordLabels.ForEach(l => WordDataGridView.Rows.Add(l, string.Empty));
    }

    private void ParseWordTable() =>
        WordDataGridView.Rows
                     .Cast<DataGridViewRow>()
                     .Take(WordDataGridView.Rows.Count - 1)
                     .ToList()
                     .ForEach(r => _keyValuePairs.Add(r.Cells[0].Value.ToString()!.Replace(" ", "_"), r.Cells[1].Value.ToString() ?? string.Empty));

    private void CreateWordDocument(string examplePath, string savePath)
    {
        using WordprocessingDocument document = WordprocessingDocument.CreateFromTemplate(examplePath);
        var body = document.MainDocumentPart!.Document.Body!;
        var keys = _keyValuePairs.Keys.ToList();
        foreach (Paragraph paragraph in body.Elements<Paragraph>())
        {
            foreach (Run run in paragraph.Elements<Run>())
            {
                foreach (Text text in run.Elements<Text>())
                {
                    for (var i = 0; i < keys.Count; i++)
                    {
                        if (text.Text.Contains(keys[i]))
                        {
                            text.Text = text.Text.Replace(keys[i], _keyValuePairs[keys[i]]);
                            keys.Remove(keys[i--]);
                        }
                    }
                }
            }
        }

        document.Clone(savePath);
    }

    private void открытьШаблонToolStripMenuItem_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new()
        {
            Filter = _filter
        };

        if (openFileDialog.ShowDialog() != DialogResult.OK)
        {
            MessageBox.Show("Не выбран шаблон", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        _examplePath = openFileDialog.FileName;
    }

    private void сохранитьРезультатToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (_examplePath == string.Empty)
        {
            MessageBox.Show("Не выбран шаблон", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        ParseWordTable();
        SaveFileDialog saveFileDialog = new()
        {
            Filter = _filter,
            FileName = "Result.docx",
        };

        if (saveFileDialog.ShowDialog() != DialogResult.OK)
        {
            MessageBox.Show("Не выбран путь для сохранения", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        CreateWordDocument(_examplePath, saveFileDialog.FileName);
    }
    #endregion
    #region Excel
    private void InitializeExcelTable() => _excelLabels
        .Take(_excelLabels.Count - 1)
        .ToList()
        .ForEach(l => ExcelDataGridView.Columns.Add(l, l));

    private bool TryParseExcelTable(out object[][] model)
    {
        model = ExcelDataGridView
            .Rows
            .Cast<DataGridViewRow>()
            .Take(ExcelDataGridView.Rows.Count - 1)
            .Select(r => TryParseRow(r, out var objects) ? objects : null)
            .ToArray()!;
        return !model.Any(x => x is null);
    }

    private bool TryParseRow(DataGridViewRow row, out object[] objects)
    {
        var cells = row.Cells;
        List<object> result = new()
        {
            row.Cells[0].Value.ToString() ?? string.Empty
        };

        for (int i = 1; i < cells.Count; i++)
        {
            if (!int.TryParse(cells[i].Value?.ToString() ?? string.Empty, out int value))
            {
                objects = Array.Empty<object>();
                return false;
            }

            result.Add(value);
        }

        objects = result.ToArray();
        return true;
    }

    private void выбратьПутьДляСохраненияToolStripMenuItem_Click(object sender, EventArgs e)
    {
        SaveFileDialog saveFileDialog = new()
        {
            Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            FileName = "Result",
        };

        if (saveFileDialog.ShowDialog() != DialogResult.OK)
        {
            MessageBox.Show("Не выбран путь для сохранения", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        _excelDocumentPath = saveFileDialog.FileName;
    }

    private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
    {
        if (!TryParseExcelTable(out var model))
        {
            MessageBox.Show("Таблица заполнена некорректно", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        CreateExcelDocument(_excelDocumentPath, model);
        MessageBox.Show("Документ создан", "Успех", MessageBoxButtons.OK, MessageBoxIcon.None);
    }

    private void CreateExcelDocument(string path, object[][] model)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet();
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        SheetData sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);
        Sheets sheets = document.WorkbookPart!.Workbook.AppendChild(new Sheets());
        Sheet sheet = new()
        {
            Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Расчеты",
        };

        sheetData.Append(CreateHeader());
        for (int i = 0; i < model.Length; i++)
        {
            var row = CreateRow(model[i], i);
            sheetData.Append(row);
        }

        sheetData.Append(CreateFooter(model.Length + 2));
        sheets.Append(sheet);
        workbookPart.Workbook.Save();
    }

    private Row CreateHeader()
    {
        Row header = new() { RowIndex = 1 };
        for (int j = 0; j < _excelLabels.Count; j++)
        {
            header.Append(new Cell()
            {
                CellReference = GetCellReference(0, j),
                CellValue = new(_excelLabels[j])
            });
        }

        return header;
    }

    private Row CreateFooter(int rowIndex)
    {
        Row footer = new() { RowIndex = Convert.ToUInt32(rowIndex) };
        footer.Append(new Cell()
        {
            CellReference = GetCellReference(rowIndex - 1, 0),
            CellValue = new("Среднее")
        });

        for (int j = 1; j < _excelLabels.Count; j++)
        {
            footer.Append(new Cell()
            {
                CellReference = GetCellReference(rowIndex - 1, j),
                CellFormula = new($"=AVERAGE({GetCellReference(1, j)}:{GetCellReference(rowIndex - 2, j)})"),
            });
        }

        return footer;
    }

    private static Row CreateRow(object[] model, int index)
    {
        Row row = new() { RowIndex = Convert.ToUInt32(index + 2) };
        for (int j = 0; j < model.Length; j++)
        {
            Cell cell = new()
            {
                CellReference = GetCellReference(index + 1, j),
                CellValue = new(model[j].ToString()!),
            };

            row.Append(cell);
        }

        Cell formulaCell = new()
        {
            CellReference = GetCellReference(index + 1, model.Length),
            CellFormula = new($"=SUM({GetCellReference(index + 1, 1)}:{GetCellReference(index + 1, model.Length - 1)})"),
        };

        row.Append(formulaCell);
        return row;
    }

    private static string GetCellReference(int i, int j) => $"{(char)('A' + j)}{i + 1}";
    #endregion
}

/*
 public void CreateStylesheet(string filePath)
 {
     using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
     {
         WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
         workbookPart.Workbook = new Workbook();
 
         WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
         stylesPart.Stylesheet = new Stylesheet();
 
         // Add cell formats
         stylesPart.Stylesheet.CellFormats = new CellFormats();
         stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
 
         // Add fonts
         stylesPart.Stylesheet.Fonts = new Fonts();
         stylesPart.Stylesheet.Fonts.AppendChild(new Font());
 
         // Add fills
         stylesPart.Stylesheet.Fills = new Fills();
         stylesPart.Stylesheet.Fills.AppendChild(new Fill());
 
         // Add borders
         stylesPart.Stylesheet.Borders = new Borders();
         stylesPart.Stylesheet.Borders.AppendChild(new Border());
 
         stylesPart.Stylesheet.Save();
     }
 }
 
 
 Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet;
 CellFormats cellformats = stylesheet.CellFormats;
 Border border = new();
 border.BottomBorder = new BottomBorder() { Style = BorderStyleValues.Thick };
 cellformats.AppendChild(border);
 */