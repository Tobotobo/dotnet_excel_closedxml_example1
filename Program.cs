using ClosedXML.Excel;

// AddressBook[] addressBooks = [
//     new(1, "東京　太郎", "tokyo_taro@example.com"),
//     new(2, "東京　花子", "tokyo_hanako@example.com"),
// ];

// using (var workbook = new XLWorkbook())
// {
//     var worksheet = workbook.Worksheets.Add("Sample Sheet");

//     // ヘッダー
//     worksheet.Cell(row: 1, column: 1).Value = "ID";
//     worksheet.Cell(row: 1, column: 2).Value = "名前";
//     worksheet.Cell(row: 1, column: 3).Value = "メールアドレス";

//     // 内容
//     var row = 2;
//     foreach(var addressBook in addressBooks) {
//         worksheet.Cell(row: row, column: 1).Value = addressBook.Id;
//         worksheet.Cell(row: row, column: 2).Value = addressBook.Name;
//         worksheet.Cell(row: row, column: 3).Value = addressBook.Mail;
//         row++;
//     }

//     workbook.SaveAs("AddressBook.xlsx");
// }


List<AddressBook> addressBooks = [];
using (var fs = new FileStream("AddressBook.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
using (var workbook = new XLWorkbook(fs))
{
    var worksheet = workbook.Worksheets.First(x => x.Name == "Sample Sheet");

    // ヘッダー
    worksheet.Cell(row: 1, column: 1).Value = "ID";
    worksheet.Cell(row: 1, column: 2).Value = "名前";
    worksheet.Cell(row: 1, column: 3).Value = "メールアドレス";

    // worksheet.CellsUsed()

    // 最終行を取得
    var lastRowNumber = worksheet.LastRowUsed().RowNumber();

    // 内容
    for(var row = 2; row <= lastRowNumber; row++) {
        var id = worksheet.Cell(row, 1).Value;
        var name = worksheet.Cell(row, 2).Value;
        var mail = worksheet.Cell(row, 3).Value;
        
        if (id.IsBlank && name.IsBlank && mail.IsBlank) {
            continue;
        }

        AddressBook addressBook;
        if (id.IsBlank) {
            // 追加
            addressBook = new AddressBook(
                0, 
                name.GetText(), 
                mail.GetText()
            );
        } else {
            // 更新
            addressBook = new AddressBook(
                (int)id.GetNumber(), 
                name.GetText(), 
                mail.GetText()
            );
        }

        addressBooks.Add(addressBook);
    }
}

foreach(var addressBook in addressBooks) {
    Console.WriteLine(addressBook);
}

