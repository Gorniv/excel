import 'dart:convert';
import 'dart:io';
import 'dart:math';
import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  test('Create New XLSX File', () {
    var excel = Excel.createExcel();
    expect(excel.sheets.entries.length, equals(1));
    expect(excel.sheets.entries.first.key, equals('Sheet1'));
  });

  test('Read XLSX File', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(excel.tables['Sheet1']!.maxColumns, equals(3));
    expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'));
  });

  test('Cell Data-Types from Microsoft Excel 365 Destkop', () {
    var file = './test/test_resources/dataTypesUsingMsExcel365Desktop.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(
      excel.tables['Tabelle1']!.rows[2][1]?.value,
      equals(TextCellValue('Some text')),
    );
    expect(
      excel.tables['Tabelle1']?.rows[3][1]?.value,
      equals(IntCellValue(42)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[4][1]?.value,
      equals(DoubleCellValue(12.3)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[5][1]?.value,
      equals(DateCellValue(year: 2023, month: 4, day: 20)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[6][1]?.value,
      equals(DateTimeCellValue(
          year: 2023, month: 4, day: 20, hour: 15, minute: 44, second: 13)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[7][1]?.value,
      equals(BoolCellValue(true)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[8][1]?.value,
      equals(BoolCellValue(false)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[9][1]?.value,
      equals(DoubleCellValue(15.99)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[10][1]?.value,
      equals(DoubleCellValue(0.05)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[11][1]?.value,
      equals(TimeCellValue(hour: 2, minute: 20, second: 10)),
    );
  });

  test('Cell Data-Types from Google Spreadsheet', () {
    var file = './test/test_resources/dataTypesUsingGoogleSpreadsheet.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(
      excel.tables['Sheet1']?.rows[2][1]?.value,
      equals(TextCellValue('Some text')),
    );
    expect(
      excel.tables['Sheet1']?.rows[3][1]?.value,
      equals(IntCellValue(42)),
    );
    expect(
      excel.tables['Sheet1']?.rows[4][1]?.value,
      equals(DoubleCellValue(12.3)),
    );
    expect(
      excel.tables['Sheet1']?.rows[5][1]?.value,
      equals(DateCellValue(year: 2023, month: 4, day: 20)),
    );
    expect(
      excel.tables['Sheet1']?.rows[6][1]?.value,
      equals(
        DateTimeCellValue(
          year: 2023,
          month: 4,
          day: 20,
          hour: 15,
          minute: 44,
          second: 13,
        ),
      ),
    );
    expect(
      excel.tables['Sheet1']?.rows[7][1]?.value,
      equals(BoolCellValue(true)),
    );
    expect(
      excel.tables['Sheet1']?.rows[8][1]?.value,
      equals(BoolCellValue(false)),
    );
    expect(
      excel.tables['Sheet1']?.rows[9][1]?.value,
      equals(DoubleCellValue(15.99)),
    );
    expect(
      excel.tables['Sheet1']?.rows[10][1]?.value,
      equals(DoubleCellValue(0.05)),
    );
  });

  test('Cell Data-Types from LibreOffice', () {
    var file = './test/test_resources/dataTypesUsingLibreoffice.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(
      excel.tables['Sheet1']?.rows[2][1]?.value,
      equals(TextCellValue('Some text')),
    );
    expect(
      excel.tables['Sheet1']?.rows[3][1]?.value,
      equals(IntCellValue(42)),
    );
    expect(
      excel.tables['Sheet1']?.rows[4][1]?.value,
      equals(DoubleCellValue(12.3)),
    );
    expect(
      excel.tables['Sheet1']?.rows[5][1]?.value,
      equals(DateCellValue(year: 2023, month: 4, day: 20)),
    );
    expect(
      excel.tables['Sheet1']?.rows[6][1]?.value,
      equals(DateTimeCellValue(
        year: 2023,
        month: 4,
        day: 20,
        hour: 15,
        minute: 44,
        second: 13,
      )),
    );
    expect(
      excel.tables['Sheet1']?.rows[7][1]?.value,
      equals(BoolCellValue(true)),
    );
    expect(
      excel.tables['Sheet1']?.rows[8][1]?.value,
      equals(BoolCellValue(false)),
    );
    expect(
      excel.tables['Sheet1']?.rows[9][1]?.value,
      equals(DoubleCellValue(15.99)),
    );
    expect(
      excel.tables['Sheet1']?.rows[10][1]?.value,
      equals(DoubleCellValue(0.05)),
    );
  });

  test('Read/Write various data types', () {
    var file = './test/test_resources/dataTypesUsingMsExcel365Desktop.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    {
      final sheet = excel.tables['Tabelle1']!;
      sheet.updateCell(
        CellIndex.indexByString('B4'),
        DoubleCellValue(13.37),
      );
      sheet.updateCell(
        CellIndex.indexByString('B5'),
        DateCellValue(year: 2025, month: 11, day: 28),
      );
      sheet.updateCell(
        CellIndex.indexByString('B6'),
        null,
      );
      sheet.updateCell(
        CellIndex.indexByString('B7'),
        TimeCellValue(hour: 20, minute: 15),
      );
      sheet.updateCell(
        CellIndex.indexByString('B8'),
        DoubleCellValue(42),
        cellStyle: CellStyle(numberFormat: NumFormat.standard_11),
      );

      final b10 = sheet.cell(CellIndex.indexByString('B10'));
      b10.cellStyle = (b10.cellStyle ?? CellStyle()).copyWith(
        numberFormat: CustomNumericNumFormat(formatCode: r'0\m\²'),
      );
    }

    final bytesWritten = excel.encode()!;
    final excelAgain = Excel.decodeBytes(bytesWritten);
    {
      final sheet = excelAgain.tables['Tabelle1']!;
      final b3 = sheet.cell(CellIndex.indexByString('B3'));
      expect(b3.value, equals(TextCellValue('Some text')));
      expect(
        b3.cellStyle?.numberFormat ?? NumFormat.standard_0,
        equals(NumFormat.standard_0),
      );

      final b4 = sheet.cell(CellIndex.indexByString('B4'));
      expect(b4.value, equals(DoubleCellValue(13.37)));
      expect(
        b4.cellStyle?.numberFormat ?? NumFormat.defaultFloat,
        equals(NumFormat.defaultFloat),
      );

      final b5 = sheet.cell(CellIndex.indexByString('B5'));
      expect(b5.value, equals(DateCellValue(year: 2025, month: 11, day: 28)));
      expect(
        b5.cellStyle?.numberFormat,
        equals(NumFormat.defaultDate),
      );

      final b6 = sheet.cell(CellIndex.indexByString('B6'));
      expect(b6.value, equals(null));
      expect(
        b6.cellStyle?.numberFormat,
        equals(NumFormat.standard_0),
      );

      final b7 = sheet.cell(CellIndex.indexByString('B7'));
      expect(b7.value, equals(TimeCellValue(hour: 20, minute: 15)));
      expect(
        b7.cellStyle?.numberFormat,
        equals(NumFormat.defaultTime),
      );

      final b8 = sheet.cell(CellIndex.indexByString('B8'));
      expect(b8.value, equals(IntCellValue(42)));
      expect(
        b8.cellStyle?.numberFormat,
        equals(NumFormat.standard_11),
      );

      final b10 = sheet.cell(CellIndex.indexByString('B10'));
      expect(b10.value, equals(DoubleCellValue(15.99)));
      expect(
        b10.cellStyle?.numberFormat,
        equals(CustomNumericNumFormat(formatCode: r'0\m\²')),
      );
    }
  });

  test('Testing customNumFormats', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];
    final format1 = CustomNumericNumFormat(formatCode: r'0.00%');
    final format2 = CustomNumericNumFormat(formatCode: r'#,##0.00');
    final styleA1 = CellStyle(
      numberFormat: format1,
    );
    final styleB1 = CellStyle(
      numberFormat: format2,
    );

    sheet.updateCell(CellIndex.indexByString('A1'), DoubleCellValue(0.15),
        cellStyle: styleA1);
    sheet.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(123456.789),
        cellStyle: styleB1);
    final bytes = excel.encode();
    final excel2 = Excel.decodeBytes(bytes!);
    final sheet2 = excel2['Sheet1'];
    final a1_2 = sheet2.cell(CellIndex.indexByString('A1'));
    final b1_2 = sheet2.cell(CellIndex.indexByString('B1'));
    expect(a1_2.cellStyle?.numberFormat, equals(format1));
    expect(a1_2.value, equals(DoubleCellValue(0.15)));
    expect(b1_2.cellStyle?.numberFormat, equals(format2));
    expect(b1_2.value, equals(DoubleCellValue(123456.789)));
  });

  group('Sheet Operations', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    Excel excel = Excel.decodeBytes(bytes);
    test('create Sheet', () {
      Sheet sheetObject = excel['SheetTmp'];
      sheetObject.insertRowIterables([
        TextCellValue('Country'),
        TextCellValue('Capital'),
        TextCellValue('Head')
      ], 0);
      sheetObject.insertRowIterables([
        TextCellValue('Russia'),
        TextCellValue('Moscow'),
        TextCellValue('Putin')
      ], 1);
      expect(excel.sheets.entries.length, equals(2));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxColumns, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(),
          equals('Putin'));
    });

    test('copy Sheet', () {
      excel.copy('SheetTmp', 'SheetTmp2');
      expect(excel.sheets.entries.length, equals(3));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxColumns, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(),
          equals('Putin'));
      expect(excel.tables['SheetTmp2']!.rows[1][2]!.value.toString(),
          equals('Putin'));
    });

    test('rename Sheet', () {
      excel.rename('SheetTmp2', 'SheetTmp3');
      expect(excel.sheets.entries.length, equals(3));
      expect(excel.tables['Sheettmp2'], equals(null));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxColumns, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(),
          equals('Putin'));
      expect(excel.tables['SheetTmp3']!.rows[1][2]!.value.toString(),
          equals('Putin'));
    });

    test('delete Sheet', () {
      excel.delete('SheetTmp3');
      excel.delete('SheetTmp');
      expect(excel.sheets.entries.length, equals(1));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
    });
  });

  test('Saving XLSX File', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    Sheet? sheetObject = excel.tables['Sheet1']!;
    sheetObject.insertRowIterables([
      TextCellValue('Russia'),
      TextCellValue('Moscow'),
      TextCellValue('Putin')
    ], 4);
    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/exampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/exampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);
    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));
    expect(newExcel.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'));
    expect(newExcel.tables['Sheet1']!.maxColumns, equals(3));
    expect(newExcel.tables['Sheet1']!.rows[4][1]!.value.toString(),
        equals('Moscow'));
  });

  test('Saving XLSX File with appendRow', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    sheet.appendRow([
      IntCellValue(8),
      DoubleCellValue(999.62221),
      DateCellValue(year: 2023, month: 4, day: 20),
      DateTimeCellValue(
        year: 2023,
        month: 4,
        day: 20,
        hour: 15,
        minute: 44,
        second: 13,
      ),
      TextCellValue('value'),
    ]);

    //stopwatch.reset();
    List<int>? fileBytes = excel.save();
    //print('saving executed in ${stopwatch.elapsed}');
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/exampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }

    var newFile = './tmp/exampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);

    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));
    expect(newExcel.tables['Sheet1']!.maxColumns, equals(5));
    expect(
        newExcel.tables['Sheet1']!.rows[0][0]!.value, equals(IntCellValue(8)));
    expect(
        newExcel.tables['Sheet1']!.rows[0][0]!.cellStyle?.numberFormat
            .toString(),
        equals(NumFormat.defaultNumeric.toString()));
    expect(newExcel.tables['Sheet1']!.rows[0][1]!.value,
        DoubleCellValue(999.62221));
    expect(
        newExcel.tables['Sheet1']!.rows[0][1]!.cellStyle?.numberFormat
            .toString(),
        equals(NumFormat.defaultFloat.toString()));
    expect(newExcel.tables['Sheet1']!.rows[0][2]!.value,
        DateCellValue(year: 2023, month: 4, day: 20));
    expect(
        newExcel.tables['Sheet1']!.rows[0][2]!.cellStyle?.numberFormat
            .toString(),
        equals(NumFormat.defaultDate.toString()));
    expect(
        newExcel.tables['Sheet1']!.rows[0][3]!.value,
        DateTimeCellValue(
          year: 2023,
          month: 4,
          day: 20,
          hour: 15,
          minute: 44,
          second: 13,
        ));
    expect(
        newExcel.tables['Sheet1']!.rows[0][3]!.cellStyle?.numberFormat
            .toString(),
        equals(NumFormat.defaultDateTime.toString()));
    expect(
        newExcel.tables['Sheet1']!.rows[0][4]!.value, TextCellValue('value'));
    expect(
        newExcel.tables['Sheet1']!.rows[0][4]!.cellStyle?.numberFormat
            .toString(),
        equals(NumFormat.standard_0.toString()));
  });

  test('Saving XLSX File with superscript', () {
    var file = './test/test_resources/superscriptExample.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/superscriptExampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/superscriptExampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);
    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));

    expect(newExcel.tables['Sheet1']!.rows[0][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[1][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[2][0]!.value.toString(),
        equals('Text in A3'));
  });

  test(
      'Add already shared strings and make sure that they are reused by checking increased usage count but equal unique count',
      () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var archive = ZipDecoder().decodeBytes(bytes);
    var sharedStringsArchive = archive.findFile('xl/sharedStrings.xml')!;

    var oldSharedStringsDocument =
        XmlDocument.parse(utf8.decode(sharedStringsArchive.content));
    var oldCount = oldSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("count");
    var oldUniqueCount = oldSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("uniqueCount");

    var excel = Excel.decodeBytes(bytes);

    Sheet? sheetObject = excel.tables['Sheet1']!;
    sheetObject.insertRowIterables([
      TextCellValue('ISRAEL'),
      TextCellValue('Jerusalem'),
      TextCellValue('Benjamin Netanyahu')
    ], 4);
    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/exampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/exampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    expect(() => Excel.decodeBytes(newFileBytes), returnsNormally);

    var newArchive = ZipDecoder().decodeBytes(newFileBytes);
    var newSharedStringsArchive = newArchive.findFile('xl/sharedStrings.xml')!;

    var newSharedStringsDocument =
        XmlDocument.parse(utf8.decode(newSharedStringsArchive.content));
    var newCount = newSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("count");
    var newUniqueCount = newSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("uniqueCount");

    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);

    expect(oldUniqueCount!.value, equals(newUniqueCount!.value));
    expect(oldCount!.value, "12");
    expect(newCount!.value, "15");
  });

  test('Saving XLSX File with superscript', () {
    var file = './test/test_resources/superscriptExample.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/superscriptExampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/superscriptExampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);
    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));

    expect(newExcel.tables['Sheet1']!.rows[0][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[1][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[2][0]!.value.toString(),
        equals('Text in A3'));
  });

  group('Header/Footer', () {
    test("Update header/footer", () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      Sheet sheetObject = excel.tables['Sheet1']!;

      sheetObject.headerFooter!.oddHeader = "Foo";
      sheetObject.headerFooter!.oddFooter = "Bar";

      var fileBytes = excel.encode();
      if (fileBytes != null) {
        File(Directory.current.path + '/tmp/exampleOut.xlsx')
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }
      var newFile = './tmp/exampleOut.xlsx';
      var newFileBytes = File(newFile).readAsBytesSync();
      var newExcel = Excel.decodeBytes(newFileBytes);
      expect(
          newExcel.tables['Sheet1']!.headerFooter!.oddHeader!, equals('Foo'));
      expect(
          newExcel.tables['Sheet1']!.headerFooter!.oddFooter!, equals('Bar'));

      // delete tmp folder only when test is successful (diagnosis)
      new Directory('./tmp').delete(recursive: true);
    });

    test("Save empty Workbook", () {
      var excel = Excel.createExcel();
      excel.save();
    });

    test("Clone header/footer of existing Workbook", () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      Sheet? sheetObject = excel.tables['Sheet1']!;

      sheetObject.headerFooter!.oddHeader = "Foo";
      sheetObject.headerFooter!.oddFooter = "Bar";

      excel.copy('Sheet1', 'test_sheet');

      Sheet? testSheet = excel.tables['test_sheet'];

      expect(testSheet!.headerFooter!.oddHeader!, equals('Foo'));
      expect(testSheet.headerFooter!.oddFooter!, equals('Bar'));
    });

    test("Remove header/footer from Workbook", () {});

    test("Reader headerFooter attributes", () {
      var file = './test/test_resources/headerFooter.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      Sheet? sheetObject = excel.tables['Sheet1']!;

      var headerFooter = sheetObject.headerFooter!;

      expect(headerFooter.alignWithMargins, isFalse);
      expect(headerFooter.differentFirst, isTrue);
      expect(headerFooter.differentOddEven, isTrue);
      expect(headerFooter.scaleWithDoc, isFalse);
    });
  });

  group('Borders', () {
    test('read file with borders', () {
      final file = './test/test_resources/borders.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;

      final borderEmpty = Border();
      final borderMedium = Border(borderStyle: BorderStyle.Medium);
      final borderMediumRed = Border(
          borderStyle: BorderStyle.Medium,
          borderColorHex: 'FFFF0000'.excelColor);
      final borderHair = Border(borderStyle: BorderStyle.Hair);
      final borderDouble = Border(borderStyle: BorderStyle.Double);

      final cellStyleA1 =
          sheetObject.cell(CellIndex.indexByString('A1')).cellStyle;
      expect(cellStyleA1?.leftBorder, equals(borderMedium));
      expect(cellStyleA1?.rightBorder, equals(borderMedium));
      expect(cellStyleA1?.topBorder, anyOf(isNull, equals(borderEmpty)));
      expect(cellStyleA1?.bottomBorder, equals(borderMediumRed));
      expect(cellStyleA1?.diagonalBorder, anyOf(isNull, equals(borderEmpty)));
      expect(cellStyleA1?.diagonalBorderUp, isFalse);
      expect(cellStyleA1?.diagonalBorderDown, isFalse);

      final cellStyleB3 =
          sheetObject.cell(CellIndex.indexByString('B3')).cellStyle;
      expect(cellStyleB3?.leftBorder, equals(borderMedium));
      expect(cellStyleB3?.rightBorder, equals(borderMedium));
      expect(cellStyleB3?.topBorder, equals(borderHair));
      expect(cellStyleB3?.bottomBorder, equals(borderHair));

      final cellStyleA5 =
          sheetObject.cell(CellIndex.indexByString('A5')).cellStyle;
      expect(cellStyleA5?.diagonalBorder, equals(borderDouble));
      expect(cellStyleA5?.diagonalBorderUp, isFalse);
      expect(cellStyleA5?.diagonalBorderDown, isTrue);

      final cellStyleC5 =
          sheetObject.cell(CellIndex.indexByString('C5')).cellStyle;
      expect(cellStyleC5?.diagonalBorder, equals(borderDouble));
      expect(cellStyleC5?.diagonalBorderUp, isTrue);
      expect(cellStyleC5?.diagonalBorderDown, isFalse);
    });

    test('test support all border styles', () {
      final file = './test/test_resources/borders2.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;

      final borderStyles = <BorderStyle>[
        BorderStyle.None,
        BorderStyle.DashDot,
        BorderStyle.DashDotDot,
        BorderStyle.Dashed,
        BorderStyle.Dotted,
        BorderStyle.Double,
        BorderStyle.Hair,
        BorderStyle.Medium,
        BorderStyle.MediumDashDot,
        BorderStyle.MediumDashDotDot,
        BorderStyle.MediumDashed,
        BorderStyle.SlantDashDot,
        BorderStyle.Thick,
        BorderStyle.Thin,
      ];

      for (var i = 1; i < borderStyles.length; ++i) {
        // Loop from i = 1, as Excel does not set None type.
        final border = Border(borderStyle: borderStyles[i]);

        final cellStyle = sheetObject
            .cell(CellIndex.indexByString('B${2 * (i + 1)}'))
            .cellStyle;

        expect(cellStyle?.leftBorder, equals(border));
        expect(cellStyle?.rightBorder, equals(border));
        expect(cellStyle?.topBorder, equals(border));
        expect(cellStyle?.bottomBorder, equals(border));
      }
    });

    test('test support for merged cells with borders', () {
      final file = './test/test_resources/mergedBorders.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;

      final borderStyles = <BorderStyle>[
        BorderStyle.None,
        BorderStyle.DashDot,
        BorderStyle.DashDotDot,
        BorderStyle.Dashed,
        BorderStyle.Dotted,
        BorderStyle.Double,
        BorderStyle.Hair,
        BorderStyle.Medium,
        BorderStyle.MediumDashDot,
        BorderStyle.MediumDashDotDot,
        BorderStyle.MediumDashed,
        BorderStyle.SlantDashDot,
        BorderStyle.Thick,
        BorderStyle.Thin,
      ];

      sheetObject.merge(
          CellIndex.indexByString('B2'), CellIndex.indexByString('D4'));

      for (var i = 1; i < borderStyles.length; ++i) {
        // Loop from i = 1, as Excel does not set None type.
        final border = Border(
            borderStyle: borderStyles[i],
            borderColorHex: "FF000000".excelColor);
        final start = CellIndex.indexByString('B${(4 * i + 2)}');
        final end = CellIndex.indexByString('D${(4 * i + 4)}');

        sheetObject.merge(start, end);

        sheetObject.setMergedCellStyle(
          start,
          CellStyle(
            leftBorder: border,
            rightBorder: border,
            topBorder: border,
            bottomBorder: border,
          ),
        );
      }

      for (var i = 1; i < borderStyles.length; ++i) {
        CellIndex cellIndexStart = CellIndex.indexByString('B${(4 * i + 2)}');
        CellIndex cellIndexEnd = CellIndex.indexByString('D${(4 * i + 4)}');

        for (var j = cellIndexStart.rowIndex; j <= cellIndexEnd.rowIndex; j++) {
          for (var k = cellIndexStart.columnIndex;
              k <= cellIndexEnd.columnIndex;
              k++) {
            final cellStyle = sheetObject
                .cell(CellIndex.indexByColumnRow(columnIndex: k, rowIndex: j))
                .cellStyle;

            final borderStyle = Border(
              borderStyle: borderStyles[i],
              borderColorHex: "FF000000".excelColor,
            );

            if (j == cellIndexStart.rowIndex) {
              expect(cellStyle?.topBorder, equals(borderStyle));
            }

            if (j == cellIndexEnd.rowIndex) {
              expect(cellStyle?.bottomBorder, equals(borderStyle));
            }

            if (k == cellIndexStart.columnIndex) {
              expect(cellStyle?.leftBorder, equals(borderStyle));
            }

            if (k == cellIndexEnd.columnIndex) {
              expect(cellStyle?.rightBorder, equals(borderStyle));
            }
          }
        }
      }
    });

    test('saving XLSX File with borders', () {
      final file = './test/test_resources/borders.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);

      final outFilePath = Directory.current.path + '/tmp/bordersOut.xlsx';
      final fileBytes = excel.encode();
      if (fileBytes != null) {
        File(outFilePath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }

      final newFileBytes = File(outFilePath).readAsBytesSync();
      final newExcel = Excel.decodeBytes(newFileBytes);
      expect(newExcel.sheets.entries.length, equals(1));

      final borderEmpty = Border();
      final borderMedium = Border(borderStyle: BorderStyle.Medium);
      final borderMediumRed = Border(
          borderStyle: BorderStyle.Medium,
          borderColorHex: 'FFFF0000'.excelColor);

      final Sheet sheetObject = newExcel.tables['Sheet1']!;
      final cellStyleB1 =
          sheetObject.cell(CellIndex.indexByString('B1')).cellStyle;
      expect(cellStyleB1?.leftBorder, equals(borderMedium));
      expect(cellStyleB1?.rightBorder, equals(borderMedium));
      expect(cellStyleB1?.topBorder, equals(borderEmpty));
      expect(cellStyleB1?.bottomBorder, equals(borderMediumRed));

      // delete tmp folder only when test is successful (diagnosis)
      new Directory('./tmp').delete(recursive: true);
    });
  });

  group('Cell Style', () {
    test('read file with rich text', () {
      final file = './test/test_resources/richText.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;
      final redHex = 'FFFF0000';
      final blueHex = 'FF2A6099';

      final cellA1 = sheetObject.cell(CellIndex.indexByString('A1')).value
          as TextCellValue;
      expect(cellA1.value.children![0].style!.fontSize, 12);
      expect(cellA1.value.children![0].style!.fontColor.colorHex, redHex);
      expect(cellA1.value.children![1].style!.fontSize, 10);
      expect(cellA1.value.children![1].style!.fontColor.colorHex, blueHex);

      final cellA2 = sheetObject.cell(CellIndex.indexByString('A2')).value
          as TextCellValue;
      expect(cellA2.value.children![0].style!.isBold, true);
      expect(cellA2.value.children![0].style!.isItalic, false);
      expect(cellA2.value.children![1].style!.isBold, false);
      expect(cellA2.value.children![1].style!.isItalic, true);

      final cellA3 = sheetObject.cell(CellIndex.indexByString('A3')).value
          as TextCellValue;
      expect(cellA3.value.children![0].style!.fontFamily, "Skia");
      expect(cellA3.value.children![1].style!.fontFamily, "Arial");
    });
  });

  group('rPh tag', () {
    test('Read Cell shared text without rPh elements', () {
      var file = './test/test_resources/rphSample.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables['Sheet1']!.rows[1][0]!.value.toString(),
          equals('plainText'));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Hellow world'));
      expect(excel.tables['Sheet1']!.rows[1][2]!.value.toString(),
          equals('世界よこんにちは'));
      expect(excel.tables['Sheet1']!.rows[2][2]!.value.toString(),
          equals('ようこそユーザー'));
      expect(excel.tables['Sheet1']!.rows[3][2]!.value.toString(),
          equals('ロケール選択'));
      expect(excel.tables['Sheet1']!.rows[4][2]!.value.toString(),
          equals('ロケール選択'));
    });

    test('saving XLSX File without rPh elements', () {
      final file = './test/test_resources/rphSample.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      excel.tables['Sheet1']!.rows[3][2]!.value = TextCellValue('ロケール選択');

      final outFilePath = Directory.current.path + '/tmp/rphSampleOut.xlsx';
      final fileBytes = excel.encode();
      if (fileBytes != null) {
        File(outFilePath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }

      final newFileBytes = File(outFilePath).readAsBytesSync();
      final newExcel = Excel.decodeBytes(newFileBytes);
      expect(newExcel.tables['Sheet1']!.rows[3][2]!.value.toString(),
          equals('ロケール選択'));

      // delete tmp folder only when test is successful (diagnosis)
      new Directory('./tmp').delete(recursive: true);
    });
  });

  group(".xls file handling", () {
    test("Exception when opening old .xls file", () {
      final file = './test/test_resources/oldXLSFile.xls';
      final bytes = File(file).readAsBytesSync();
      try {
        Excel.decodeBytes(bytes);
      } catch (e) {
        expect(e, isA<UnsupportedError>());
        expect(
            e.toString(),
            equals(
                'Unsupported operation: Excel format unsupported. Only .xlsx files are supported'));
      }
    });

    test("Exception when opening new .xls file", () {
      final file = './test/test_resources/newXLSFile.xls';
      final bytes = File(file).readAsBytesSync();
      try {
        Excel.decodeBytes(bytes);
      } catch (e) {
        expect(e, isA<UnsupportedError>());
        expect(
            e.toString(),
            equals(
                'Unsupported operation: Excel format unsupported. Only .xlsx files are supported'));
      }
    });

    test('Sheet Remove and Rename Operations', () {
      final List<Excel> excelFiles =
          List<Excel>.generate(5, (_) => Excel.createExcel());

      final List<List<int>> data = List<List<int>>.generate(
          5, (x) => List<int>.generate(5, (i) => (x + 1) * (i + 1)));

      const newName = 'Sheet1Replacement';

      const defaultSheetName = 'Sheet1';

      final backgroundColor =
          ExcelColor.values.where((e) => e.type == ColorType.material).toList();
      final fontColor =
          ExcelColor.values.where((e) => e.type == ColorType.color).toList();
      final borderColor = ExcelColor.values
          .where((e) => e.type == ColorType.materialAccent)
          .toList();

      excelFiles.forEach((element) {
        expect(element.getDefaultSheet()!, defaultSheetName);
        for (var row = 0; row < data.length; row++) {
          for (var column = 0; column < data[row].length; column++) {
            final border = Border(
              borderColorHex: borderColor[column],
              borderStyle: BorderStyle.Thin,
            );

            element.updateCell(
              element.getDefaultSheet()!,
              CellIndex.indexByColumnRow(columnIndex: column, rowIndex: row),
              IntCellValue(data[row][column]),
              cellStyle: CellStyle()
                ..bottomBorder = border
                ..topBorder = border
                ..leftBorder = border
                ..rightBorder = border
                ..backgroundColor = backgroundColor[row]
                ..fontColor = fontColor[column],
            );
          }
        }

        if (Random().nextBool()) {
          /// Rename test
          element.rename(element.getDefaultSheet()!, newName);
          expect(element.getDefaultSheet(), null);
          element.setDefaultSheet(newName);
          expect(element.getDefaultSheet(), newName);
        } else {
          /// Remove test
          element.copy(element.getDefaultSheet()!, newName);
          expect(element.getDefaultSheet()!, defaultSheetName);
          element.delete(element.getDefaultSheet()!);
          expect(element.getDefaultSheet(), null);
          element.setDefaultSheet(newName);
          expect(element.getDefaultSheet()!, newName);
        }

        expect(element.tables.length, 1);

        for (var row = 0; row < data.length; row++) {
          for (var column = 0; column < data[row].length; column++) {
            var cell = element.tables[newName]?.rows[row][column];
            expect(cell?.cellStyle?.backgroundColor, backgroundColor[row]);
            expect(cell?.cellStyle?.fontColor, fontColor[column]);
            expect([
              cell?.cellStyle?.bottomBorder.borderColorHex,
              cell?.cellStyle?.topBorder.borderColorHex,
              cell?.cellStyle?.leftBorder.borderColorHex,
              cell?.cellStyle?.rightBorder.borderColorHex,
            ], everyElement(borderColor[column].colorHex));
          }
        }
      });
    });
  });

  group('Spanned Items', () {
    test("read spanned items", () {
      var file = './test/test_resources/spannedItemExample.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      Sheet? sheet = excel.tables["Spanned Items"]!;

      testSpannedItemsSheetValues(Sheet sheet) {
        final cells =
            sheet.rows.expand((r) => r.where((c) => c != null)).toList();

        expect(cells[0]?.value, equals(TextCellValue('spanned item A1:B1')));
        expect(cells[0]?.cellIndex,
            equals(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)));

        expect(cells[1]?.value, equals(TextCellValue('spanned item A2:A3')));
        expect(cells[1]?.cellIndex,
            equals(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1)));

        expect(cells[2]?.value, equals(TextCellValue('spanned item A4:B5')));
        expect(cells[2]?.cellIndex,
            equals(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 3)));
      }

      testSpannedItemsList(Sheet sheet) {
        List<String> spannedItems = sheet.spannedItems;

        expect(spannedItems[0], equals('A1:B1'));
        expect(spannedItems[1], equals('A2:A3'));
        expect(spannedItems[2], equals('A4:B5'));
      }

      testSpannedItemsList(sheet);

      testSpannedItemsSheetValues(sheet);

      var fileBytes = excel.encode();
      if (fileBytes != null) {
        File(Directory.current.path + '/tmp/spannedItemExampleOut.xlsx')
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }
      var newFile = './tmp/spannedItemExampleOut.xlsx';
      var newFileBytes = File(newFile).readAsBytesSync();
      var newExcel = Excel.decodeBytes(newFileBytes);
      // delete tmp folder
      new Directory('./tmp').delete(recursive: true);

      Sheet? newSheet = newExcel.tables["Spanned Items"]!;

      testSpannedItemsList(newSheet);

      testSpannedItemsSheetValues(newSheet);
    });
  });

  test('Parse column width row height', () {
    var file = './test/test_resources/columnWidthRowHeight.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    Sheet? sheetObject = excel.tables['Sheet1']!;

    // should 20 with a litle bit of tolerance.
    expect(sheetObject.defaultColumnWidth, greaterThan(18));
    expect(sheetObject.defaultColumnWidth, lessThan(22));

    // should 20 with a litle bit of tolerance.
    expect(sheetObject.defaultRowHeight, greaterThan(18));
    expect(sheetObject.defaultRowHeight, lessThan(22));

    // should 40 with a litle bit of tolerance.
    expect(sheetObject.getColumnWidth(1), greaterThan(38));
    expect(sheetObject.getColumnWidth(1), lessThan(42));

    // should 40 with a litle bit of tolerance.
    expect(sheetObject.getRowHeight(1), greaterThan(38));
    expect(sheetObject.getRowHeight(1), lessThan(42));
  });

  test('Decode customNumFmtIdBelow164.xlsx without throwing exception', () {
    var file = './test/test_resources/customNumFmtIdBelow164.xlsx';

    expect(
      () {
        final bytes = File(file).readAsBytesSync();
        final _ = Excel.decodeBytes(bytes);
      },
      returnsNormally,
      reason: 'Decoding the file should not throw any exception',
    );
  });

  test('Saving XLSX File with max width calc', () {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];
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

- Interactive Cross-Localization Table:
Maximize global visibility by validating and optimizing metadata across 60 countries. Tailored suggestions ensure strategic keyword deployment.

- Subscription Pricing Management:
Seamlessly manage subscriptions, localize descriptions, and optimize performance with effortless A/B testing.
Optimize your app’s global revenue strategy by managing prices, applying multipliers (Netflix, Big Mac, PPP), apply VAT by country.

- Bulk upload hundreds of screenshots and videos effortlessly with automatic device and locale detection and perfect sizing.

- Custom Product Pages (CPP):
Tailor app pages for targeted marketing and A/B testing. Easily copy and create customized CPPs with localized content, videos, and screenshots.

- In-App Events:
Boost user engagement and visibility through compelling events aligned with global holidays and key marketing dates. Localize events effortlessly for maximum impact.
Visualize trending events and gain market insights through interactive heatmaps and detailed event analytics.

- Unlimited ASO:
Enjoy limitless keyword and project management capabilities for unparalleled optimization.

- ASO Check:
Instantly evaluate app performance across all App Store locales, identify optimization opportunities, and enhance global reach.

- Comprehensive App Info:
Gain deep insights into keyword rankings, reviews, competitor strategies, and app metrics. Analyze detailed performance data to stay ahead of competition.

- Keywords Ranking & Competitor Spy:
Track unlimited keywords, analyze competitor strategies, and export comprehensive keyword data for strategic ASO decisions.

- Keyword Lists & Competitor Analysis:
Efficiently organize, track, and manage keywords and competitor apps. Identify opportunities to differentiate and outperform the competition.

- Top Apps & Category Ranking:
Analyze top-performing apps across various categories, platforms, and countries, monitoring trends and competitor positioning effectively.

- User Reviews Management:
Easily manage, translate, and respond to user reviews across languages. Use AI-driven responses and swiftly report inappropriate content.

- Metrics & Analytics:
Monitor app performance with detailed metrics, including impressions, views, revenue, and user engagement. Make data-driven decisions effortlessly.

- Streamline your localization process, translating content instantly across multiple languages and formats with AI-powered tools.

- Project Sharing & Collaboration:
Collaborate effectively with flexible access controls, enabling efficient teamwork and client management.

- Timeline & Historical Insights:
Track key app changes and metadata updates over time, gaining insights for improved future strategies. Historical data is immediately available.

- iTunes Country Switcher

Join thousands of developers and marketers who trust ASO.dev to simplify their ASO workflow and drive app growth.
Terms and Conditions - https://aso.dev/terms/"
''';
    sheet.appendRow([
      TextCellValue('1 test',
          cellStyle: CellStyle(textWrapping: TextWrapping.WrapText)),
    ]);
    sheet.appendRow([
      TextCellValue(bigText,
          cellStyle: CellStyle(textWrapping: TextWrapping.WrapText)),
    ]);
    sheet.setColumnAutoFit(0);

    //stopwatch.reset();
    List<int>? fileBytes = excel.save();
    //print('saving executed in ${stopwatch.elapsed}');
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/exampleOut_max.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }

    var newFile = './tmp/exampleOut_max.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);

    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    final widths = newExcel.tables['Sheet1']!.getColumnWidth(0);
    print('widths: $widths');
    expect(widths > 10.0 && widths < 230.0, isTrue);
  });
}
