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
  final filenameRegExp = RegExp(
    r'([^-])-(\d\d)(\d\d)([.]([^ ]+)| )([^.]+).xlsx',
  );
  final DateTime start;
  if (now.weekday > 1) {
    start = now.add(Duration(days: 8 - now.weekday));
  } else {
    start = now;
  }
  final baseDirectory = mastersDirectory.parent;
  for (var i = 0; i < 2; i++) {
    final DateTime date;
    if (i == 0) {
      date = start;
    } else {
      date = start.add(Duration(days: i * 7));
    }
    final weekNumber = (date.day / 7).floor() + 1;
    final year = date.year;
    final month = date.month.toString().padLeft(2, '0');
    final day = date.day.toString().padLeft(2, '0');
    final yearDirectory = Directory(
      path.join(baseDirectory.path, 'Booking Sheets $year'),
    );
    final weekDirectory = Directory(
      path.join(yearDirectory.path, 'Booking Sheets (wc $year-$month-$day)'),
    );
    if (!weekDirectory.existsSync()) {
      weekDirectory.createSync(recursive: true);
    }
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
        if (a1 == null) {
          throw StateError('A1 should not be blank, not $a1.');
        }
        final DateTime groupStartDate;
        final groupDay = int.parse(dayGroup);
        if (groupDay == 1) {
          // Start date always falls on a Monday.
          groupStartDate = date;
        } else {
          groupStartDate = date.add(Duration(days: groupDay - 1));
        }
        sheet.updateCell(
          a1.cellIndex,
          DateCellValue(
            year: groupStartDate.year,
            month: groupStartDate.month,
            day: groupStartDate.day,
          ),
          cellStyle: a1.cellStyle,
        );
        destination.writeAsBytesSync(excel.save()!);
      }
    }
  }
}
