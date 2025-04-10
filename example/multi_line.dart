import 'dart:io';

import 'package:excel/excel.dart';

void main() {
  final excel = Excel.createExcel();

  const defaultSheetName = 'Sheet1';
  final sheet = excel.sheets[defaultSheetName];

  var data = List.generate(
      10,
      (c) => TextCellValue('$c \r\n 2',
          cellStyle: CellStyle(textWrapping: TextWrapping.WrapText)));
  for (var i = 0; i < 10; i++) {
    sheet!.appendRow(data);
  }
  final bytes = excel.encode();
  if (bytes != null) {
    File('example/example.xlsx')
      ..createSync()
      ..writeAsBytesSync(bytes);
  }
}
