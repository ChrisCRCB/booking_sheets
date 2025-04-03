import 'dart:io';

import 'package:excel/excel.dart';
import 'package:path/path.dart' as path;

void main(final List<String> arguments) {
  final Directory mastersDirectory;
  if (arguments.length == 1) {
    mastersDirectory = Directory(arguments.single);
  } else {
    mastersDirectory = Directory.current;
  }
  final now = DateTime.now();
  final DateTime startDate;
  if (now.weekday > 1) {
    startDate = now.add(Duration(days: 8 - now.weekday));
  } else {
    startDate = now;
  }
  final weekNumber = (startDate.day / 7).floor() + 1;
  final baseDirectory = mastersDirectory.parent;
  final year = startDate.year;
  final month = startDate.month.toString().padLeft(2, '0');
  final day = startDate.day.toString().padLeft(2, '0');
  final yearDirectory = Directory(
    path.join(baseDirectory.path, 'Booking Sheets $year'),
  );
  final weekDirectory = Directory(
    path.join(yearDirectory.path, 'Booking Sheets (wc $year-$month-$day)'),
  );
  if (!weekDirectory.existsSync()) {
    weekDirectory.createSync(recursive: true);
  }
  final filenameRegExp = RegExp(
    r'([^-])-(\d\d)(\d\d)([.]([^ ]+)| )([^.]+).xlsx',
  );
  for (final file in mastersDirectory.listSync().whereType<File>()) {
    final basename = path.basename(file.path);
    final match = filenameRegExp.firstMatch(basename);
    if (match == null) {
      continue;
    }
    final dayGroup = match.group(1)!;
    final weekGroup = match.group(5);
    if (weekGroup != null) {
      final groupWeekNumber = int.parse(weekGroup);
      if (weekNumber != groupWeekNumber) {
        continue;
      }
    }
    final destination = File(path.join(weekDirectory.path, basename));
    if (!destination.existsSync()) {
      final bytes = file.readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final sheet = excel.sheets[excel.sheets.keys.first];
      final a1 = sheet!.rows.first.first;
      if (a1 != null) {
        throw StateError('A1 should be blank, not $a1.');
      }
      final DateTime groupStartDate;
      final groupDay = int.parse(dayGroup);
      if (groupDay == 1) {
        // Start date always falls on a Monday.
        groupStartDate = startDate;
      } else {
        groupStartDate = startDate.add(Duration(days: groupDay - 1));
      }
      sheet.updateCell(
        CellIndex.indexByString('a1'),
        DateCellValue(
          year: groupStartDate.year,
          month: groupStartDate.month,
          day: groupStartDate.day,
        ),
      );
      destination.writeAsBytesSync(excel.save()!);
    }
  }
}
