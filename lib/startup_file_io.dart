import 'dart:io';
import 'package:cross_file/cross_file.dart';

XFile? xFileFromArgs(List<String> args) {
  if (args.isEmpty) return null;
  final path = args.first;
  if (!File(path).existsSync()) return null;
  return XFile(path);
}