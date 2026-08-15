import 'dart:typed_data';
import 'package:file_picker/file_picker.dart';

Future<Uint8List> readFileBytes(PlatformFile file) async {
  return await file.xFile.readAsBytes();
}