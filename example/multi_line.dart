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
  for (var i = 0; i < 1; i++) {
    sheet!.appendRow(data);
  }
  final bigText = '''
"ASO.dev – The Ultimate Alternative Client for App Store Connect with comprehensive App Store Optimization (ASO) Tools, empowering your app's growth. Automate your app release process seamlessly, saving valuable time and ensuring accuracy for every update.

Your API keys and sensitive data are securely stored on your device.
Optimize app metadata securely and easily via official App Store Connect API.
Requests to App Store Connect are made directly from your device—no server-side requests, ensuring maximum security.
Manage security flexibly with individual or team-based API keys.

- Smart Metadata Editor:
Create unique, validated metadata with real-time keyword analytics, competitor insights, and seamless rollback options. Simplify macOS and iOS releases with instant metadata copying.

- Bulk Editor:
Rapidly manage updates across multiple localizations. Instantly translate content using AI (OpenAI, Claude, DeepSeek), Google Translate, or DeepL. Easily detect duplicates and compare previous versions.

Join thousands of developers and marketers who trust ASO.dev to simplify their ASO workflow and drive app growth.
Terms and Conditions - https://aso.dev/terms/"
''';
  sheet!.setColumnAutoFit(0);
  final newBigData = TextCellValue(bigText,
      cellStyle: CellStyle(textWrapping: TextWrapping.WrapText));
  sheet.appendRow([newBigData]);
  final bytes = excel.encode();
  if (bytes != null) {
    File('example/test/multi_line.xlsx')
      ..createSync()
      ..writeAsBytesSync(bytes);
  }
}
