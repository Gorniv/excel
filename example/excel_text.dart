import 'dart:io';

import 'package:excel/excel.dart';

void main() {
  final excel = Excel.createExcel();

  const defaultSheetName = 'Sheet1';
  final sheet = excel.sheets[defaultSheetName];
  final text = '''
Dear HWilliams26,\n\nThank you for sharing your thoughts with us. We're sorry to hear that our features didn’t meet your expectations. Our goal is to offer the best solution for controlling your TV from a smartphone, alongside a great user experience. We genuinely value the feedback and support from our users as we work to improve. If you have any additional questions, feel free to reach out to our support team at contact.tvsmart@gmail.com. We’d be happy to assist you further!\n\n\bSincerely,\nContact TVSmart Support Team
''';
  var data = TextCellValue('$text',
      cellStyle: CellStyle(textWrapping: TextWrapping.WrapText));
  sheet!.appendRow([data]);
  final bytes = excel.encode();
  if (bytes != null) {
    File('example/text.xlsx')
      ..createSync()
      ..writeAsBytesSync(bytes);
  }
}
