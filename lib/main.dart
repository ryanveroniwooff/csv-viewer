import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:file_picker/file_picker.dart';
import 'package:csv/csv.dart';
import 'package:excel/excel.dart' as xlsx;
import 'package:archive/archive.dart';
import 'file_bytes_reader.dart';
import 'dart:convert';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    final colorScheme = const ColorScheme(
      brightness: Brightness.dark,
      primary: Color(0xFF4C7CF3),
      onPrimary: Colors.white,
      secondary: Color(0xFF6D8CC7),
      onSecondary: Colors.white,
      error: Color(0xFFCF6679),
      onError: Colors.black,
      surface: Color(0xFF1A1D24),
      onSurface: Color(0xFFE3E5E8),
      surfaceContainerLowest: Color(0xFF121419),
      surfaceContainerLow: Color(0xFF1F232B),
      surfaceContainer: Color(0xFF242832),
      surfaceContainerHigh: Color(0xFF2B303B),
      surfaceContainerHighest: Color(0xFF333944),
      onSurfaceVariant: Color(0xFFA3A9B4),
      outline: Color(0xFF474D59),
    );

    return MaterialApp(
      title: 'RowMaster',
      themeMode: ThemeMode.dark,
      darkTheme: ThemeData(
        useMaterial3: true,
        colorScheme: colorScheme,
        scaffoldBackgroundColor: colorScheme.surface,
        appBarTheme: AppBarTheme(
          backgroundColor: colorScheme.surface,
          surfaceTintColor: Colors.transparent,
          elevation: 0,
          scrolledUnderElevation: 0,
        ),
        cardTheme: CardThemeData(
          elevation: 0,
          color: colorScheme.surfaceContainer,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(16),
          ),
        ),
        inputDecorationTheme: InputDecorationTheme(
          filled: true,
          fillColor: colorScheme.surfaceContainerHighest,
          border: OutlineInputBorder(
            borderRadius: BorderRadius.circular(12),
            borderSide: BorderSide.none,
          ),
          contentPadding:
              const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
        ),
        dialogTheme: DialogThemeData(
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(20),
          ),
        ),
      ),
      home: const CsvViewerPage(),
    );
  }
}

class FilterCondition {
  String column;
  String operatorType;
  String value;

  FilterCondition({
    required this.column,
    required this.operatorType,
    required this.value,
  });

  String get operatorLabel {
    switch (operatorType) {
      case 'contains':
        return 'contains';
      case 'equals':
        return '=';
      case 'not_equals':
        return '≠';
      case '>':
        return '>';
      case '<':
        return '<';
      default:
        return operatorType;
    }
  }
}

class CsvViewerPage extends StatefulWidget {
  const CsvViewerPage({super.key});

  @override
  State<CsvViewerPage> createState() => _CsvViewerPageState();
}

class _CsvViewerPageState extends State<CsvViewerPage> {
  List<List<dynamic>>? _rows;
  String? _fileName;
  String? _error;

  String _searchQuery = '';
  String? _groupByColumn;
  final List<FilterCondition> _filters = [];
  final Map<String, List<String>> _uniqueValueCache = {};
  List<double> _columnWidths = [];
  final Set<String> _hiddenColumns = {};

  List<dynamic>? _editingRow;
  int? _editingCol;
  final TextEditingController _editController = TextEditingController();
  final FocusNode _editFocusNode = FocusNode();

  @override
  void initState() {
    super.initState();
    _editFocusNode.addListener(() {
      if (!_editFocusNode.hasFocus && _editingRow != null) {
        _commitEdit();
      }
    });
  }

  @override
  void dispose() {
    _editController.dispose();
    _editFocusNode.dispose();
    super.dispose();
  }

  List<dynamic> get _headers => _rows != null ? _rows!.first : [];
  List<List<dynamic>> get _dataRows =>
      _rows != null ? _rows!.skip(1).toList() : [];

  List<int> get _visibleColumnIndices => List.generate(_headers.length, (i) => i)
      .where((i) => !_hiddenColumns.contains(_headers[i].toString()))
      .toList();

  void _hideColumn(String column) {
    setState(() => _hiddenColumns.add(column));
  }

  void _showColumn(String column) {
    setState(() => _hiddenColumns.remove(column));
  }

  void _startEdit(List<dynamic> row, int col) {
    setState(() {
      _editingRow = row;
      _editingCol = col;
      _editController.text = row[col].toString();
      _editController.selection = TextSelection(
        baseOffset: 0,
        extentOffset: _editController.text.length,
      );
    });
    WidgetsBinding.instance
        .addPostFrameCallback((_) => _editFocusNode.requestFocus());
  }

  void _commitEdit() {
    final row = _editingRow;
    final col = _editingCol;
    if (row == null || col == null) return;

    setState(() {
      row[col] = _editController.text;
      _editingRow = null;
      _editingCol = null;
      _uniqueValueCache.clear();
    });
  }

  void _cancelEdit() {
    setState(() {
      _editingRow = null;
      _editingCol = null;
    });
    _editFocusNode.unfocus();
  }

  List<String> _uniqueValuesForColumn(String column) {
    if (_uniqueValueCache.containsKey(column)) {
      return _uniqueValueCache[column]!;
    }
    final colIndex = _headers.indexOf(column);
    if (colIndex == -1) return [];
    final values = _dataRows
        .map((row) => row[colIndex].toString())
        .where((v) => v.isNotEmpty)
        .toSet()
        .toList();
    values.sort();
    _uniqueValueCache[column] = values;
    return values;
  }

  void _computeColumnWidths() {
    final widths = <double>[];
    final sample = _dataRows.take(200);
    for (var c = 0; c < _headers.length; c++) {
      var maxLen = _headers[c].toString().length.toDouble();
      for (final row in sample) {
        if (c < row.length) {
          final len = row[c].toString().length.toDouble();
          if (len > maxLen) maxLen = len;
        }
      }
      widths.add((maxLen * 9 + 32).clamp(100, 320));
    }
    _columnWidths = widths;
  }

  dynamic _extractCellPrimitive(xlsx.CellValue? cellValue) {
    if (cellValue == null) return '';
    try {
      final dynamic raw = (cellValue as dynamic).value;
      if (raw != null) return raw;
    } catch (_) {
    }
    return cellValue.toString();
  }

  Future<Uint8List> _sanitizeXlsxBytes(Uint8List bytes) async {
    final archive = ZipDecoder().decodeBytes(bytes);
    final emptyStringCellPattern = RegExp(r'<c([^>]*)></c>');

    final newArchive = Archive();
    for (final file in archive.files) {
      if (file.isFile &&
          file.name.startsWith('xl/worksheets/') &&
          file.name.endsWith('.xml')) {
        final content = utf8.decode(file.content as List<int>);
        final fixed = content.replaceAllMapped(emptyStringCellPattern, (m) {
          final attrs = m.group(1)!;
          if (attrs.contains('t="s"')) {
            final cleaned = attrs.replaceAll(RegExp(r'\s*t="s"'), '');
            return '<c$cleaned></c>';
          }
          return m.group(0)!;
        });
        final fixedBytes = utf8.encode(fixed);
        newArchive.addFile(
          ArchiveFile(file.name, fixedBytes.length, fixedBytes),
        );
      } else {
        newArchive.addFile(file);
      }
    }

    final encoded = ZipEncoder().encode(newArchive);
    if (encoded == null) {
      throw Exception('Failed to re-encode sanitized xlsx');
    }
    return Uint8List.fromList(encoded);
  }

  Future<List<List<dynamic>>> _parseXlsx(PlatformFile file) async {
    final rawBytes = await readFileBytes(file);

    Uint8List bytes;
    try {
      bytes = await _sanitizeXlsxBytes(rawBytes);
    } catch (_) {
      bytes = rawBytes;
    }

    final workbook = xlsx.Excel.decodeBytes(bytes);
    if (workbook.tables.isEmpty) return [];

    final sheetName = workbook.tables.keys.first;
    final sheet = workbook.tables[sheetName]!;

    return sheet.rows
        .map((row) =>
            row.map((cell) => _extractCellPrimitive(cell?.value)).toList())
        .toList();
  }

  Future<List<List<dynamic>>> _parseFile(PlatformFile file) async {
    if (file.name.toLowerCase().endsWith('.xlsx')) {
      return _parseXlsx(file);
    }
    final bytes = await readFileBytes(file);
    return Csv().decode(utf8.decode(bytes));
  }

  Future<void> _pickFile() async {
    setState(() => _error = null);

    final files = await FilePicker.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['csv', 'xlsx'],
      withData: true,
    );

    if (files.isEmpty) return;

    try {
      final parsed = await _parseFile(files.first);

      setState(() {
        _rows = parsed;
        _fileName = files.first.name;
        _searchQuery = '';
        _filters.clear();
        _groupByColumn = null;
        _hiddenColumns.clear();
        _uniqueValueCache.clear();
        _computeColumnWidths();
      });
    } catch (e) {
      setState(() => _error = 'Could not read file: $e');
    }
  }

  bool _headersMatch(List<dynamic> a, List<dynamic> b) {
    if (a.length != b.length) return false;
    for (var i = 0; i < a.length; i++) {
      if (a[i].toString().trim().toLowerCase() !=
          b[i].toString().trim().toLowerCase()) {
        return false;
      }
    }
    return true;
  }

  Future<void> _mergeFiles() async {
    if (_rows == null) return;
    setState(() => _error = null);

    final files = await FilePicker.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['csv', 'xlsx'],
      withData: true,
    );

    if (files.isEmpty) return;

    final rowsToAdd = <List<dynamic>>[];
    final skipped = <String>[];
    var mergedFileCount = 0;

    for (final file in files) {
      try {
        final parsed = await _parseFile(file);

        if (parsed.isEmpty) {
          skipped.add('${file.name} (empty file)');
          continue;
        }

        final fileHeaders = parsed.first;
        if (!_headersMatch(fileHeaders, _headers)) {
          skipped.add('${file.name} (headers don\'t match)');
          continue;
        }

        rowsToAdd.addAll(parsed.skip(1));
        mergedFileCount++;
      } catch (e) {
        skipped.add('${file.name} (error reading file)');
      }
    }

    if (rowsToAdd.isNotEmpty) {
      setState(() {
        _rows = [_headers, ..._dataRows, ...rowsToAdd];
        _uniqueValueCache.clear();
        _computeColumnWidths();
      });
    }

    if (!mounted) return;

    final messageParts = <String>[];
    if (mergedFileCount > 0) {
      messageParts.add(
          'Merged ${rowsToAdd.length} rows from $mergedFileCount file${mergedFileCount == 1 ? '' : 's'}');
    }
    if (skipped.isNotEmpty) {
      messageParts.add('Skipped: ${skipped.join(', ')}');
    }

    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(messageParts.join('. ')),
        duration: const Duration(seconds: 5),
      ),
    );
  }

  List<List<dynamic>> get _filteredRows {
    var rows = _dataRows;

    if (_searchQuery.isNotEmpty) {
      final q = _searchQuery.toLowerCase();
      rows = rows
          .where((row) =>
              row.any((cell) => cell.toString().toLowerCase().contains(q)))
          .toList();
    }

    for (final filter in _filters) {
      if (filter.value.isEmpty) continue;
      final colIndex = _headers.indexOf(filter.column);
      if (colIndex == -1) continue;

      rows = rows.where((row) {
        final cellValue = row[colIndex].toString();
        switch (filter.operatorType) {
          case 'contains':
            return cellValue.toLowerCase().contains(filter.value.toLowerCase());
          case 'equals':
            return cellValue.toLowerCase() == filter.value.toLowerCase();
          case 'not_equals':
            return cellValue.toLowerCase() != filter.value.toLowerCase();
          case '>':
            final a = num.tryParse(cellValue);
            final b = num.tryParse(filter.value);
            if (a == null || b == null) return false;
            return a > b;
          case '<':
            final a = num.tryParse(cellValue);
            final b = num.tryParse(filter.value);
            if (a == null || b == null) return false;
            return a < b;
          default:
            return true;
        }
      }).toList();
    }

    return rows;
  }

  List<MapEntry<String, int>> get _groupedData {
    if (_groupByColumn == null) return [];
    final colIndex = _headers.indexOf(_groupByColumn);
    if (colIndex == -1) return [];

    final counts = <String, int>{};
    for (final row in _filteredRows) {
      final key = colIndex < row.length ? row[colIndex].toString() : '';
      counts[key] = (counts[key] ?? 0) + 1;
    }

    final entries = counts.entries.toList()
      ..sort((a, b) {
        final byCount = b.value.compareTo(a.value);
        if (byCount != 0) return byCount;
        return a.key.compareTo(b.key);
      });
    return entries;
  }

  ({List<List<dynamic>> rows, String baseName, int count}) _currentExportData() {
    if (_groupByColumn != null) {
      return (
        rows: [
          [_groupByColumn, 'Count'],
          ..._groupedData.map((e) => [e.key, e.value]),
        ],
        baseName: 'export_grouped',
        count: _groupedData.length,
      );
    }
    final visible = _visibleColumnIndices;
    final visibleHeaders = visible.map((i) => _headers[i]).toList();
    final visibleFilteredRows = _filteredRows
        .map((row) => visible.map((i) => i < row.length ? row[i] : '').toList())
        .toList();
    return (
      rows: [visibleHeaders, ...visibleFilteredRows],
      baseName: 'export',
      count: visibleFilteredRows.length,
    );
  }

  xlsx.CellValue _toCellValue(dynamic value) {
    if (value is int) return xlsx.IntCellValue(value);
    if (value is double) return xlsx.DoubleCellValue(value);
    if (value is bool) return xlsx.BoolCellValue(value);
    return xlsx.TextCellValue(value?.toString() ?? '');
  }

  Future<void> _exportCsv() async {
    if (_rows == null) return;
    final data = _currentExportData();

    final csvString = Csv().encode(data.rows);
    final bytes = Uint8List.fromList(utf8.encode(csvString));

    try {
      final savedPath = await FilePicker.saveFile(
        dialogTitle: 'Export CSV',
        fileName: '${data.baseName}.csv',
        bytes: bytes,
      );

      if (!mounted) return;

      if (savedPath != null) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Exported ${data.count} rows')),
        );
      }
    } catch (e) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Export failed: $e')),
      );
    }
  }

  Future<void> _exportXlsx() async {
    if (_rows == null) return;
    final data = _currentExportData();

    try {
      final workbook = xlsx.Excel.createExcel();
      final sheet = workbook['Sheet1'];
      for (final row in data.rows) {
        sheet.appendRow(row.map(_toCellValue).toList());
      }

      final encoded = workbook.save();
      if (encoded == null) {
        throw Exception('Failed to encode xlsx');
      }

      final savedPath = await FilePicker.saveFile(
        dialogTitle: 'Export Excel',
        fileName: '${data.baseName}.xlsx',
        bytes: Uint8List.fromList(encoded),
      );

      if (!mounted) return;

      if (savedPath != null) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Exported ${data.count} rows')),
        );
      }
    } catch (e) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Export failed: $e')),
      );
    }
  }

  void _openFilterDialog() {
    showDialog(
      context: context,
      builder: (context) {
        final colorScheme = Theme.of(context).colorScheme;
        return StatefulBuilder(
          builder: (context, setDialogState) {
            return AlertDialog(
              title: const Text('Filters'),
              content: SizedBox(
                width: 560,
                child: SingleChildScrollView(
                  child: Column(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      ..._filters.asMap().entries.map((entry) {
                        final index = entry.key;
                        final filter = entry.value;
                        final suggestions =
                            _uniqueValuesForColumn(filter.column);

                        return Padding(
                          padding: const EdgeInsets.symmetric(vertical: 6.0),
                          child: Row(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Expanded(
                                flex: 3,
                                child: DropdownButtonFormField<String>(
                                  initialValue: filter.column,
                                  isExpanded: true,
                                  items: _headers
                                      .map((h) => DropdownMenuItem(
                                            value: h.toString(),
                                            child: Text(
                                              h.toString(),
                                              overflow: TextOverflow.ellipsis,
                                            ),
                                          ))
                                      .toList(),
                                  onChanged: (v) {
                                    setDialogState(() {
                                      filter.column = v!;
                                      filter.value = '';
                                    });
                                  },
                                ),
                              ),
                              const SizedBox(width: 8),
                              Expanded(
                                flex: 2,
                                child: DropdownButtonFormField<String>(
                                  initialValue: filter.operatorType,
                                  isExpanded: true,
                                  items: const [
                                    DropdownMenuItem(
                                        value: 'contains',
                                        child: Text('contains')),
                                    DropdownMenuItem(
                                        value: 'equals', child: Text('=')),
                                    DropdownMenuItem(
                                        value: 'not_equals',
                                        child: Text('≠')),
                                    DropdownMenuItem(
                                        value: '>', child: Text('>')),
                                    DropdownMenuItem(
                                        value: '<', child: Text('<')),
                                  ],
                                  onChanged: (v) {
                                    setDialogState(
                                        () => filter.operatorType = v!);
                                  },
                                ),
                              ),
                              const SizedBox(width: 8),
                              Expanded(
                                flex: 3,
                                child: Autocomplete<String>(
                                  key: ValueKey('filter_value_$index'),
                                  initialValue:
                                      TextEditingValue(text: filter.value),
                                  optionsBuilder: (textEditingValue) {
                                    if (textEditingValue.text.isEmpty) {
                                      return suggestions.take(50);
                                    }
                                    return suggestions
                                        .where((v) => v
                                            .toLowerCase()
                                            .contains(textEditingValue.text
                                                .toLowerCase()))
                                        .take(50);
                                  },
                                  onSelected: (selection) {
                                    setDialogState(
                                        () => filter.value = selection);
                                  },
                                  fieldViewBuilder: (context, controller,
                                      focusNode, onFieldSubmitted) {
                                    return TextFormField(
                                      controller: controller,
                                      focusNode: focusNode,
                                      decoration: const InputDecoration(
                                        hintText: 'value',
                                      ),
                                      onChanged: (v) => filter.value = v,
                                    );
                                  },
                                  optionsViewBuilder:
                                      (context, onSelected, options) {
                                    return Align(
                                      alignment: Alignment.topLeft,
                                      child: Material(
                                        elevation: 4,
                                        borderRadius:
                                            BorderRadius.circular(12),
                                        color:
                                            colorScheme.surfaceContainerHigh,
                                        child: ConstrainedBox(
                                          constraints: const BoxConstraints(
                                              maxHeight: 220, maxWidth: 220),
                                          child: ListView.builder(
                                            padding: EdgeInsets.zero,
                                            shrinkWrap: true,
                                            itemCount: options.length,
                                            itemBuilder: (context, i) {
                                              final option =
                                                  options.elementAt(i);
                                              return ListTile(
                                                dense: true,
                                                title: Text(option),
                                                onTap: () =>
                                                    onSelected(option),
                                              );
                                            },
                                          ),
                                        ),
                                      ),
                                    );
                                  },
                                ),
                              ),
                              IconButton(
                                icon: const Icon(Icons.delete_outline),
                                onPressed: () {
                                  setDialogState(() {
                                    _filters.removeAt(index);
                                  });
                                },
                              ),
                            ],
                          ),
                        );
                      }),
                      const SizedBox(height: 4),
                      Align(
                        alignment: Alignment.centerLeft,
                        child: TextButton.icon(
                          onPressed: _headers.isEmpty
                              ? null
                              : () {
                                  setDialogState(() {
                                    _filters.add(FilterCondition(
                                      column: _headers.first.toString(),
                                      operatorType: 'contains',
                                      value: '',
                                    ));
                                  });
                                },
                          icon: const Icon(Icons.add),
                          label: const Text('Add filter'),
                        ),
                      ),
                    ],
                  ),
                ),
              ),
              actions: [
                TextButton(
                  onPressed: () {
                    setDialogState(() => _filters.clear());
                  },
                  child: const Text('Clear all'),
                ),
                FilledButton(
                  onPressed: () {
                    setState(() {});
                    Navigator.pop(context);
                  },
                  child: const Text('Apply'),
                ),
              ],
            );
          },
        );
      },
    );
  }

  @override
  Widget build(BuildContext context) {
    final colorScheme = Theme.of(context).colorScheme;

    return Scaffold(
      appBar: AppBar(
        title: Row(
          children: [
            Icon(Icons.table_chart_outlined, color: colorScheme.primary),
            const SizedBox(width: 10),
            Text(_fileName ?? 'RowMaster'),
          ],
        ),
      ),
      body: _rows == null ? _buildEmptyState(colorScheme) : _buildLoaded(colorScheme),
    );
  }

  Widget _buildEmptyState(ColorScheme colorScheme) {
    return Center(
      child: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          Icon(Icons.insert_drive_file_outlined,
              size: 64, color: colorScheme.onSurfaceVariant),
          const SizedBox(height: 16),
          Text(
            'No file loaded',
            style: Theme.of(context)
                .textTheme
                .titleMedium
                ?.copyWith(color: colorScheme.onSurfaceVariant),
          ),
          const SizedBox(height: 20),
          FilledButton.icon(
            onPressed: _pickFile,
            icon: const Icon(Icons.folder_open),
            label: const Text('Browse for CSV or Excel'),
          ),
          if (_error != null) ...[
            const SizedBox(height: 16),
            Text(_error!, style: TextStyle(color: colorScheme.error)),
          ],
        ],
      ),
    );
  }

  Widget _buildLoaded(ColorScheme colorScheme) {
    final filtered = _filteredRows;
    return Padding(
      padding: const EdgeInsets.all(20.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          SizedBox(
            height: 48,
            child: Row(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                FilledButton.tonalIcon(
                  style: FilledButton.styleFrom(
                    minimumSize: const Size(0, 48),
                  ),
                  onPressed: _pickFile,
                  icon: const Icon(Icons.folder_open),
                  label: const Text('Browse'),
                ),
                const SizedBox(width: 12),
                OutlinedButton.icon(
                  style: OutlinedButton.styleFrom(
                    minimumSize: const Size(0, 48),
                  ),
                  onPressed: _mergeFiles,
                  icon: const Icon(Icons.merge_type),
                  label: const Text('Merge'),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: TextField(
                    decoration: const InputDecoration(
                      prefixIcon: Icon(Icons.search),
                      hintText: 'Search all fields...',
                      isDense: true,
                      contentPadding: EdgeInsets.symmetric(vertical: 14),
                    ),
                    onChanged: (v) => setState(() => _searchQuery = v),
                  ),
                ),
                const SizedBox(width: 12),
                Badge(
                  label: Text('${_filters.length}'),
                  isLabelVisible: _filters.isNotEmpty,
                  child: OutlinedButton.icon(
                    style: OutlinedButton.styleFrom(
                      minimumSize: const Size(0, 48),
                    ),
                    onPressed: _openFilterDialog,
                    icon: const Icon(Icons.filter_list),
                    label: const Text('Filters'),
                  ),
                ),
                const SizedBox(width: 12),
                SizedBox(
                  width: 190,
                  child: DropdownButtonFormField<String?>(
                    initialValue: _groupByColumn,
                    isExpanded: true,
                    decoration: const InputDecoration(
                      prefixIcon: Icon(Icons.workspaces_outlined),
                      isDense: true,
                      contentPadding: EdgeInsets.symmetric(vertical: 14),
                    ),
                    hint: const Text('Group by'),
                    items: [
                      const DropdownMenuItem<String?>(
                        value: null,
                        child: Text('None'),
                      ),
                      ..._headers.map(
                        (h) => DropdownMenuItem<String?>(
                          value: h.toString(),
                          child: Text(
                            h.toString(),
                            overflow: TextOverflow.ellipsis,
                          ),
                        ),
                      ),
                    ],
                    onChanged: (v) => setState(() => _groupByColumn = v),
                  ),
                ),
                const SizedBox(width: 12),
                PopupMenuButton<String>(
                  tooltip: 'Export',
                  onSelected: (value) {
                    if (value == 'csv') _exportCsv();
                    if (value == 'xlsx') _exportXlsx();
                  },
                  itemBuilder: (context) => const [
                    PopupMenuItem(
                      value: 'csv',
                      child: Row(
                        children: [
                          Icon(Icons.description_outlined, size: 18),
                          SizedBox(width: 10),
                          Text('Export as CSV'),
                        ],
                      ),
                    ),
                    PopupMenuItem(
                      value: 'xlsx',
                      child: Row(
                        children: [
                          Icon(Icons.grid_on_outlined, size: 18),
                          SizedBox(width: 10),
                          Text('Export as Excel (.xlsx)'),
                        ],
                      ),
                    ),
                  ],
                  child: Container(
                    height: 48,
                    padding: const EdgeInsets.symmetric(horizontal: 16),
                    decoration: BoxDecoration(
                      color: colorScheme.secondaryContainer,
                      borderRadius: BorderRadius.circular(24),
                    ),
                    child: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        Icon(Icons.file_download_outlined,
                            size: 18, color: colorScheme.onSecondaryContainer),
                        const SizedBox(width: 8),
                        Text(
                          'Export',
                          style: TextStyle(
                              color: colorScheme.onSecondaryContainer),
                        ),
                        const SizedBox(width: 2),
                        Icon(Icons.arrow_drop_down,
                            size: 18, color: colorScheme.onSecondaryContainer),
                      ],
                    ),
                  ),
                ),
              ],
            ),
          ),
          if (_filters.isNotEmpty) ...[
            const SizedBox(height: 12),
            Wrap(
              spacing: 8,
              runSpacing: 8,
              children: _filters
                  .map(
                    (f) => InputChip(
                      label: Text('${f.column} ${f.operatorLabel} "${f.value}"'),
                      onDeleted: () => setState(() => _filters.remove(f)),
                    ),
                  )
                  .toList(),
            ),
          ],
          if (_hiddenColumns.isNotEmpty) ...[
            const SizedBox(height: 12),
            Wrap(
              spacing: 8,
              runSpacing: 8,
              children: [
                Padding(
                  padding: const EdgeInsets.only(top: 6),
                  child: Text(
                    'Hidden:',
                    style: Theme.of(context)
                        .textTheme
                        .bodySmall
                        ?.copyWith(color: colorScheme.onSurfaceVariant),
                  ),
                ),
                ..._hiddenColumns.map(
                  (col) => ActionChip(
                    avatar: const Icon(Icons.visibility_outlined, size: 16),
                    label: Text(col),
                    onPressed: () => _showColumn(col),
                  ),
                ),
              ],
            ),
          ],
          if (_error != null) ...[
            const SizedBox(height: 12),
            Text(_error!, style: TextStyle(color: colorScheme.error)),
          ],
          const SizedBox(height: 14),
          Text(
            _groupByColumn == null
                ? '${filtered.length} of ${_dataRows.length} rows'
                : '${_groupedData.length} unique values in "$_groupByColumn" '
                    '(from ${filtered.length} filtered rows)',
            style: Theme.of(context)
                .textTheme
                .bodySmall
                ?.copyWith(color: colorScheme.onSurfaceVariant),
          ),
          const SizedBox(height: 10),
          Expanded(
            child: _groupByColumn == null
                ? _buildTable(colorScheme, filtered)
                : _buildGroupTable(colorScheme),
          ),
        ],
      ),
    );
  }

  Widget _buildTable(ColorScheme colorScheme, List<List<dynamic>> rows) {
    if (_columnWidths.length != _headers.length) {
      _computeColumnWidths();
    }
    final visible = _visibleColumnIndices;
    final totalWidth =
        visible.fold<double>(0, (a, c) => a + _widthForColumn(c));

    return Card(
      clipBehavior: Clip.antiAlias,
      child: SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: SizedBox(
          width: totalWidth,
          child: Column(
            children: [
              _buildHeaderRow(colorScheme),
              Divider(height: 1, thickness: 1, color: colorScheme.outline),
              Expanded(
                child: ListView.builder(
                  itemExtent: 44,
                  itemCount: rows.length,
                  itemBuilder: (context, i) {
                    final row = rows[i];
                    return Container(
                      color: i.isEven
                          ? Colors.transparent
                          : colorScheme.surfaceContainerLow,
                      child: Row(
                        children: visible.map((c) {
                          final width = _widthForColumn(c);
                          final isEditing =
                              identical(row, _editingRow) && _editingCol == c;

                          if (isEditing) {
                            return SizedBox(
                              width: width,
                              child: Padding(
                                padding:
                                    const EdgeInsets.symmetric(horizontal: 4),
                                child: Focus(
                                  onKeyEvent: (node, event) {
                                    if (event is KeyDownEvent &&
                                        event.logicalKey ==
                                            LogicalKeyboardKey.escape) {
                                      _cancelEdit();
                                      return KeyEventResult.handled;
                                    }
                                    return KeyEventResult.ignored;
                                  },
                                  child: TextField(
                                    controller: _editController,
                                    focusNode: _editFocusNode,
                                    autofocus: true,
                                    style: const TextStyle(fontSize: 14),
                                    decoration: const InputDecoration(
                                      isDense: true,
                                      contentPadding: EdgeInsets.symmetric(
                                          horizontal: 8, vertical: 8),
                                      border: OutlineInputBorder(),
                                    ),
                                    onSubmitted: (_) => _commitEdit(),
                                  ),
                                ),
                              ),
                            );
                          }

                          return GestureDetector(
                            behavior: HitTestBehavior.opaque,
                            onTap: () => _startEdit(row, c),
                            child: Container(
                              width: width,
                              padding: const EdgeInsets.symmetric(
                                  horizontal: 16),
                              alignment: Alignment.centerLeft,
                              child: Text(
                                c < row.length ? row[c].toString() : '',
                                overflow: TextOverflow.ellipsis,
                                maxLines: 1,
                              ),
                            ),
                          );
                        }).toList(),
                      ),
                    );
                  },
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  double _widthForColumn(int c) =>
      c < _columnWidths.length ? _columnWidths[c] : 140.0;

  Widget _buildGroupTable(ColorScheme colorScheme) {
    final data = _groupedData;
    const valueWidth = 320.0;
    const countWidth = 120.0;

    return Card(
      clipBehavior: Clip.antiAlias,
      child: Column(
        children: [
          Container(
            height: 48,
            color: colorScheme.surfaceContainerHigh,
            child: Row(
              children: [
                Container(
                  width: valueWidth,
                  padding: const EdgeInsets.symmetric(horizontal: 16),
                  alignment: Alignment.centerLeft,
                  child: Text(
                    _groupByColumn ?? '',
                    style: const TextStyle(fontWeight: FontWeight.w600),
                    overflow: TextOverflow.ellipsis,
                  ),
                ),
                Container(
                  width: countWidth,
                  padding: const EdgeInsets.symmetric(horizontal: 16),
                  alignment: Alignment.centerLeft,
                  child: const Text(
                    'Count',
                    style: TextStyle(fontWeight: FontWeight.w600),
                  ),
                ),
              ],
            ),
          ),
          Divider(height: 1, thickness: 1, color: colorScheme.outline),
          Expanded(
            child: ListView.builder(
              itemExtent: 44,
              itemCount: data.length,
              itemBuilder: (context, i) {
                final entry = data[i];
                return Container(
                  color: i.isEven
                      ? Colors.transparent
                      : colorScheme.surfaceContainerLow,
                  child: Row(
                    children: [
                      Container(
                        width: valueWidth,
                        padding: const EdgeInsets.symmetric(horizontal: 16),
                        alignment: Alignment.centerLeft,
                        child: Text(
                          entry.key,
                          overflow: TextOverflow.ellipsis,
                          maxLines: 1,
                        ),
                      ),
                      Container(
                        width: countWidth,
                        padding: const EdgeInsets.symmetric(horizontal: 16),
                        alignment: Alignment.centerLeft,
                        child: Text('${entry.value}'),
                      ),
                    ],
                  ),
                );
              },
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildHeaderRow(ColorScheme colorScheme) {
    return Container(
      height: 48,
      color: colorScheme.surfaceContainerHigh,
      child: Row(
        children: _visibleColumnIndices.map((c) {
          final width = _widthForColumn(c);
          final header = _headers[c].toString();
          return Container(
            width: width,
            padding: const EdgeInsets.only(left: 16, right: 8),
            alignment: Alignment.centerLeft,
            child: Row(
              children: [
                Expanded(
                  child: Text(
                    header,
                    style: const TextStyle(fontWeight: FontWeight.w600),
                    overflow: TextOverflow.ellipsis,
                    maxLines: 1,
                  ),
                ),
                Tooltip(
                  message: 'Hide column',
                  child: InkWell(
                    borderRadius: BorderRadius.circular(4),
                    onTap: () => _hideColumn(header),
                    child: Padding(
                      padding: const EdgeInsets.all(3.0),
                      child: Icon(
                        Icons.visibility_off_outlined,
                        size: 15,
                        color: colorScheme.onSurfaceVariant,
                      ),
                    ),
                  ),
                ),
              ],
            ),
          );
        }).toList(),
      ),
    );
  }
}