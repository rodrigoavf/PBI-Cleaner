from dataclasses import dataclass
from typing import Sequence, Callable, Iterable
import re

from PyQt6.QtCore import Qt, QEvent, QRegularExpression, QObject
from PyQt6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont, QTextCursor
from PyQt6.QtWidgets import QCompleter


# A list of common DAX keywords and functions for highlighting and completion
DAX_KEYWORDS = [
    # Structural / Query Keywords
    'DEFINE', 'MEASURE', 'EVALUATE', 'VAR', 'RETURN', 'TABLE', 'COLUMN',
    'ROW', 'SELECT', 'FROM', 'WHERE', 'ORDER BY', 'GROUP BY',
    # Logical / Comparison / Operators (keywords that appear as terms)
    'TRUE', 'FALSE', 'BLANK', 'AND', 'OR', 'NOT', 'IN',
    # Conditional / branching
    'IF', 'THEN', 'ELSE', 'SWITCH',
    # Iteration / context
    'FOR', 'TO', 'BY', 'AS',
    # Modeling / Data construction
    'DATATABLE', 'SUMMARIZECOLUMNS', 'GENERATE', 'GENERATEALL',
    # Query-specific
    'START AT', 'DEFINE', 'MEASURE', 'EVALUATE', 'ORDER BY',
]

DAX_FUNCTIONS = [
    # Aggregation / Basic
    'SUM', 'AVERAGE', 'AVERAGEX', 'COUNT', 'COUNTA', 'COUNTAX', 'COUNTROWS',
    'MIN', 'MINA', 'MAX', 'MAXA', 'PRODUCT', 'PRODUCTX', 'MEDIANX',
    'GEOMEANX', 'PERCENTILEX.EXC', 'PERCENTILEX.INC', 'CONCATENATEX',
    'RANKX',

    # Iterator functions (row by row)
    'SUMX', 'AVERAGEX', 'COUNTX', 'MAXX', 'MINX', 'PRODUCTX',
    'FILTERX', 'ADDCOLUMNS', 'SELECTCOLUMNS', 'GENERATE', 'GENERATEALL',
    'GENERATESERIES', 'CROSSJOIN', 'UNION', 'INTERSECT', 'EXCEPT',
    'NATURALINNERJOIN', 'NATURALLEFTOUTERJOIN', 'GROUPBY', 'CURRENTGROUP',
    'ROLLUP', 'ROLLUPADDISSUBTOTAL', 'ROLLUPGROUP', 'ROLLUPISSUBTOTAL',

    # Filter / Context / Lookup
    'CALCULATE', 'CALCULATETABLE', 'FILTER', 'ALL', 'ALLSELECTED',
    'ALLEXCEPT', 'REMOVEFILTERS', 'KEEPFILTERS', 'VALUES', 'DISTINCT',
    'CROSSFILTER', 'USERELATIONSHIP', 'RELATED', 'RELATEDTABLE',
    'LOOKUPVALUE', 'EARLIER', 'EARLIEST', 'ISINSCOPE', 'HASONEFILTER',
    'HASONEVALUE', 'TREATAS', 'ISFILTERED', 'ISCROSSFILTERED', 'NAMEOF',
    'SELECTEDMEASURE',

    # Time Intelligence
    'DATEADD', 'DATESINPERIOD', 'DATESBETWEEN', 'DATESYTD', 'DATESMTD',
    'DATESQTD', 'TOTALYTD', 'TOTALMTD', 'TOTALQTD', 'SAMEPERIODLASTYEAR',
    'PARALLELPERIOD', 'PREVIOUSYEAR', 'NEXTYEAR', 'PREVIOUSMONTH',
    'NEXTMONTH', 'PREVIOUSDAY', 'NEXTDAY', 'STARTOFYEAR', 'ENDOFYEAR',
    'STARTOFMONTH', 'ENDOFMONTH', 'STARTOFQUARTER', 'ENDOFQUARTER',
    'CLOSINGBALANCEWEEK', 'CLOSINGBALANCEMONTH', 'CLOSINGBALANCEQUARTER',
    'CLOSINGBALANCEYEAR', 'DATESWTD', 'DATESYTD', 'DATESQTD',

    # Date & Time / Conversion
    'DATE', 'DATEDIFF', 'DATEVALUE', 'DAY', 'EDATE', 'EOMONTH',
    'HOUR', 'MINUTE', 'MONTH', 'NETWORKDAYS', 'NOW', 'QUARTER',
    'SECOND', 'TIME', 'TIMEVALUE', 'TODAY', 'UTCNOW', 'UTCTODAY',
    'WEEKDAY', 'WEEKNUM', 'YEAR', 'YEARFRAC',

    # Text / String functions
    'COMBINEVALUES', 'CONCATENATE', 'CONCATENATEX', 'EXACT', 'FIND',
    'FIXED', 'FORMAT', 'LEFT', 'LEN', 'LOWER', 'MID', 'REPLACE', 'REPT',
    'RIGHT', 'SEARCH', 'SUBSTITUTE', 'UPPER', 'UNICHAR', 'UNICODE',

    # Mathematical & Trigonometric
    'ABS', 'CEILING', 'FLOOR', 'EXP', 'LN', 'LOG', 'LOG10', 'MOD',
    'POWER', 'QUOTIENT', 'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'SIGN', 'SQRT',
    'SQRTPI', 'DIVIDE',

    # Logical / Information
    'AND', 'OR', 'NOT', 'IF', 'IFERROR', 'COALESCE', 'ISBLANK', 'ISNUMBER',
    'ISTEXT', 'ISEMPTY', 'ISNONTEXT', 'ERROR', 'SWITCH', 'SELECTEDVALUE',

    # Parent / Child / Hierarchy
    'PATH', 'PATHCONTAINS', 'PATHITEM', 'PATHITEMREVERSE', 'PATHLENGTH',

    # Statistical / Distribution (less common)
    'NORM.DIST', 'NORM.S.DIST', 'NORM.INV', 'NORM.S.INV', 'BETA.DIST', 'BETA.INV',
    'BINOM.DIST', 'CHISQ.DIST', 'CHISQ.INV', 'CHISQ.TEST', 'CONFIDENCE.NORM',
    'EXPON.DIST', 'F.DIST', 'F.INV', 'F.TEST', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV',
    'LOGNORM.DIST', 'LOGNORM.INV', 'POISSON.DIST', 'T.DIST', 'T.INV', 'T.TEST',

    # Utility / Conversion / Others
    'BLANK', 'VALUE', 'INT', 'TRUNC', 'CURRENCY', 'EXACT', 'CONVERT',
    'ISO.CEILING',

    # Metadata / Info functions
    'USERNAME', 'USERPRINCIPALNAME', 'CUSTOMDATA', 'INFO.VIEW.COLUMNS', 'INFO.VIEW.TABLES', 'INFO.VIEW.MEASURES',
    'INFO.VIEW.RELATIONSHIPS', 
    
    # Query / Table output
    'DATATABLE', 'SUMMARIZE', 'SUMMARIZECOLUMNS',
    'TOPN', 'RANK', 'ROWNUMBER', 'MATCHBY', 'LOOKUPWITHTOTALS', 'FIRST', 'LAST', 'NEXT', 'PREVIOUS',
]

DAX_MODEL_TABLES: list[str] = []
DAX_MODEL_COLUMNS: list[str] = []


def _dax_model_terms() -> list[str]:
    return DAX_MODEL_TABLES + DAX_MODEL_COLUMNS

M_KEYWORDS = [
    # Core syntax
    'let', 'in', 'each', 'if', 'then', 'else', 'try', 'otherwise', 'error',
    'and', 'or', 'not', 'is', 'as', 'meta', 'section', 'shared',
    'true', 'false', 'null',

    # Types
    'type', 'any', 'nullable', 'function', 'table', 'list', 'record',
    'number', 'text', 'logical', 'date', 'time', 'datetime', 'datetimezone', 'duration',

    # Miscellaneous
    'optional', 'binary', 'none', 'anynonnull', 'number', 'text', 'logical',
    'each', 'value', 'key', 'try', 'error', 'metadata',
]

M_FUNCTIONS = [
    # --- Table functions ---
    'Table.AddColumn', 'Table.RemoveColumns', 'Table.SelectColumns', 'Table.SelectRows',
    'Table.ExpandTableColumn', 'Table.ExpandRecordColumn', 'Table.Combine', 'Table.Distinct',
    'Table.TransformColumnTypes', 'Table.NestedJoin', 'Table.Join', 'Table.FromRecords',
    'Table.FromList', 'Table.RenameColumns', 'Table.Group', 'Table.Sort',
    'Table.PromoteHeaders', 'Table.AddIndexColumn', 'Table.RemoveRows', 'Table.FirstN',
    'Table.LastN', 'Table.Skip', 'Table.FillDown', 'Table.FillUp',
    'Table.ReplaceValue', 'Table.Transpose', 'Table.Unpivot', 'Table.UnpivotOtherColumns',
    'Table.Buffer', 'Table.Column', 'Table.ColumnCount', 'Table.ColumnNames', 'Table.RowCount',
    'Table.SelectRowsWithErrors', 'Table.RemoveRowsWithErrors', 'Table.TransformColumns',
    'Table.CombineColumns', 'Table.SplitColumn', 'Table.DemoteHeaders', 'Table.FromColumns',
    'Table.FromRows', 'Table.ToRows', 'Table.ToColumns', 'Table.ToRecords', 'Table.ReorderColumns',
    'Table.AddKey', 'Table.HasColumns', 'Table.Schema', 'Table.Partition', 'Table.ReplaceRows',

    # --- List functions ---
    'List.Transform', 'List.Accumulate', 'List.Generate', 'List.Zip',
    'List.FirstN', 'List.Skip', 'List.Sort', 'List.Distinct',
    'List.Sum', 'List.Average', 'List.Max', 'List.Min', 'List.RemoveNulls',
    'List.Contains', 'List.ContainsAny', 'List.ContainsAll',
    'List.Combine', 'List.PositionOf', 'List.PositionOfAny', 'List.FindText',
    'List.Reverse', 'List.RemoveItems', 'List.RemoveMatchingItems',
    'List.InsertRange', 'List.RemoveRange', 'List.ReplaceValue', 'List.Select',
    'List.First', 'List.Last', 'List.Count', 'List.Repeat', 'List.Dates', 'List.Times',
    'List.Numbers', 'List.Accumulate', 'List.TransformMany', 'List.Buffer', 'List.Split',

    # --- Record functions ---
    'Record.Field', 'Record.AddField', 'Record.RemoveFields', 'Record.ToTable',
    'Record.FromList', 'Record.FromTable', 'Record.FieldNames', 'Record.FieldValues',
    'Record.Combine', 'Record.HasFields', 'Record.SelectFields', 'Record.RenameFields',
    'Record.RemoveFields', 'Record.ReorderFields', 'Record.TransformFields',

    # --- Text functions ---
    'Text.Upper', 'Text.Lower', 'Text.Trim', 'Text.Length', 'Text.Combine', 'Text.Split', 'Text.Replace',
    'Text.Start', 'Text.End', 'Text.Middle', 'Text.PositionOf', 'Text.PositionOfAny', 'Text.Contains',
    'Text.StartsWith', 'Text.EndsWith', 'Text.BeforeDelimiter', 'Text.AfterDelimiter',
    'Text.PadStart', 'Text.PadEnd', 'Text.Remove', 'Text.RemoveRange', 'Text.ToList', 'Text.FromBinary',
    'Text.Proper', 'Text.NewGuid', 'Text.Repeat', 'Text.BetweenDelimiters', 'Text.Format', 'Text.Select',
    'Text.SplitAny', 'Text.CombineWithDelimiter', 'Text.From', 'Text.ToBinary',

    # --- Number functions ---
    'Number.From', 'Number.ToText', 'Number.Round', 'Number.RoundDown', 'Number.RoundUp',
    'Number.Abs', 'Number.Power', 'Number.Mod', 'Number.Sqrt', 'Number.Log', 'Number.Exp',
    'Number.Sign', 'Number.Random', 'Number.RandomBetween', 'Number.Max', 'Number.Min',
    'Number.IsNaN', 'Number.IsEven', 'Number.IsOdd', 'Number.IntegerDivide',

    # --- Date and Time functions ---
    'DateTime.LocalNow', 'DateTimeZone.FixedUtcNow', 'Date.AddDays', 'Date.From',
    'DateTime.From', 'DateTime.AddZone', 'DateTime.ToText', 'DateTime.FromText',
    'DateTimeZone.SwitchZone', 'DateTimeZone.RemoveZone', 'DateTimeZone.ToLocal',
    'Date.StartOfMonth', 'Date.EndOfMonth', 'Date.Day', 'Date.Month', 'Date.Year',
    'Date.AddMonths', 'Date.AddYears', 'Date.AddQuarters', 'Date.AddWeeks',
    'Time.From', 'Time.ToText', 'Time.FromText', 'Time.Hour', 'Time.Minute', 'Time.Second',

    # --- Logical / Value functions ---
    'Value.Is', 'Value.Type', 'Value.ReplaceType', 'Value.Metadata', 'Value.RemoveMetadata',
    'Value.ReplaceMetadata', 'Value.As', 'Value.Equals', 'Value.Compare', 'Value.Add', 'Value.Subtract',
    'Value.Divide', 'Value.Multiply', 'Value.NullableEquals', 'Value.NonNullEquals',
    'Logical.FromText', 'Logical.From', 'Logical.ToText',

    # --- Binary functions ---
    'Binary.Decompress', 'Binary.Buffer', 'Binary.FromList', 'Binary.FromText', 'Binary.ToText',
    'Binary.Compress', 'Binary.Combine', 'Binary.Format', 'Binary.Length', 'Binary.Reverse',
    'Binary.From', 'Binary.ToList',

    # --- Function helpers ---
    'Function.From', 'Function.Invoke', 'Function.InvokeAfter', 'Function.FromText',
    'Function.IsDataSource', 'Function.IsDeterministic',

    # --- Error handling ---
    'Error.Record', 'Error.Reason', 'Error.Message', 'Error.Detail',

    # --- Type and metadata ---
    'Type.Is', 'Type.ForRecord', 'Type.AddTableKey', 'Type.AddTablePrimaryKey', 'Type.AddMetadata',

    # --- Environment / evaluation ---
    'Expression.Evaluate', 'Expression.Identifier', 'Expression.Constant', 'Expression.TryEvaluate',
    'Environment.Name', 'Environment.Set', 'Environment.Value',

    # --- Miscellaneous ---
    'Uri.Parts', 'Uri.BuildQueryString', 'Uri.Combine',
    'Json.Document', 'Csv.Document', 'Xml.Document', 'Binary.Decompress',
    'Excel.CurrentWorkbook', 'Excel.Workbook', 'File.Contents', 'Folder.Files', 'Folder.Contents',
    'Web.Contents', 'Web.Page', 'Web.BrowserContents', 'SharePoint.Files', 'SharePoint.Tables',
    'PowerPlatform.Dataflows', 'Odbc.DataSource', 'Sql.Database', 'MySql.Database',
    'PostgreSQL.Database', 'Oracle.Database', 'OleDb.DataSource',
    'AzureStorage.BlobContents', 'AzureStorage.Contents',
    'Table.Profile', 'Table.MatchesAllRows', 'Table.MatchesAnyRows', 'Table.First',
    'Table.Last', 'Table.Range', 'Table.AddRankColumn', 'Table.View',
    'Record.ToList', 'Record.TransformFields', 'List.TransformMany', 'List.Accumulate',
]

class DAXHighlighter(QSyntaxHighlighter):
    """Simple DAX syntax highlighter for QTextDocument."""

    def __init__(self, document):
        super().__init__(document)

        # Formats
        self.f_keyword = QTextCharFormat()
        self.f_keyword.setForeground(QColor('#C586C0'))  # purple-ish
        self.f_keyword.setFontWeight(QFont.Weight.DemiBold)

        self.f_function = QTextCharFormat()
        self.f_function.setForeground(QColor('#4FC1FF'))  # blue

        self.f_string = QTextCharFormat()
        self.f_string.setForeground(QColor('#CE9178'))  # orange

        self.f_number = QTextCharFormat()
        self.f_number.setForeground(QColor('#B5CEA8'))  # green

        self.f_comment = QTextCharFormat()
        self.f_comment.setForeground(QColor('#6A9955'))  # comment green

        # Build regex rules (use inline (?i) for case-insensitive)
        # Keywords (word-boundary)
        kw_pattern = r"(?i)\b(" + "|".join(DAX_KEYWORDS) + r")\b"
        self.re_keyword = QRegularExpression(kw_pattern)

        # Functions followed by optional whitespace and '(' (highlight the name)
        fn_pattern = r"(?i)\b(" + "|".join(DAX_FUNCTIONS) + r")\b(?=\s*\()"
        self.re_function = QRegularExpression(fn_pattern)

        # Strings: double quotes
        self.re_string = QRegularExpression(r'"[^"\\]*(?:\\.[^"\\]*)*"')

        # Numbers (ints/floats)
        self.re_number = QRegularExpression(r"\b[0-9]+(\.[0-9]+)?\b")

        # Line comments //... and --...
        self.re_line_comment = QRegularExpression(r"//.*$")
        self.re_line_comment_dash = QRegularExpression(r"--.*$")

        # Block comments /* ... */ (multi-line handled in highlightBlock via setCurrentBlockState)
        self.re_comment_start = QRegularExpression(r"/\*")
        self.re_comment_end = QRegularExpression(r"\*/")

    def highlightBlock(self, text: str) -> None:
        # Block comment handling
        self.setCurrentBlockState(0)

        comment_spans: list[tuple[int, int]] = []

        text_len = len(text)
        if self.previousBlockState() == 1:
            end_index = text.find("*/")
            if end_index == -1:
                comment_spans.append((0, text_len))
                self.setCurrentBlockState(1)
                start_index = -1
            else:
                end_index += 2
                comment_spans.append((0, end_index))
                start_index = text.find("/*", end_index)
        else:
            start_index = text.find("/*")

        while start_index != -1:
            end_index = text.find("*/", start_index + 2)
            if end_index == -1:
                comment_spans.append((start_index, text_len - start_index))
                self.setCurrentBlockState(1)
                break
            end_index += 2
            comment_spans.append((start_index, end_index - start_index))
            start_index = text.find("/*", end_index)

        # Single line comment
        it = self.re_line_comment.globalMatch(text)
        while it.hasNext():
            m = it.next()
            comment_spans.append((m.capturedStart(), m.capturedLength()))

        it = self.re_line_comment_dash.globalMatch(text)
        while it.hasNext():
            m = it.next()
            comment_spans.append((m.capturedStart(), m.capturedLength()))

        # Strings
        string_spans: list[tuple[int, int]] = []

        it = self.re_string.globalMatch(text)
        while it.hasNext():
            m = it.next()
            start = m.capturedStart()
            length = m.capturedLength()
            if self._span_overlaps(comment_spans, start, length):
                continue
            string_spans.append((start, length))
            if start > 0 and text[start - 1] == '#':
                continue
            self.setFormat(start, length, self.f_string)

        # Numbers
        it = self.re_number.globalMatch(text)
        while it.hasNext():
            m = it.next()
            if not self._span_overlaps(comment_spans, m.capturedStart(), m.capturedLength()):
                self.setFormat(m.capturedStart(), m.capturedLength(), self.f_number)

        # Functions
        it = self.re_function.globalMatch(text)
        while it.hasNext():
            m = it.next()
            if not self._span_overlaps(comment_spans, m.capturedStart(1), m.capturedLength(1)):
                self.setFormat(m.capturedStart(1), m.capturedLength(1), self.f_function)

        # Keywords
        it = self.re_keyword.globalMatch(text)
        while it.hasNext():
            m = it.next()
            if not self._span_overlaps(comment_spans, m.capturedStart(1), m.capturedLength(1)):
                self.setFormat(m.capturedStart(1), m.capturedLength(1), self.f_keyword)

        # Apply comment formatting last so it overrides other spans
        for start, length in comment_spans:
            if length > 0:
                self.setFormat(start, length, self.f_comment)

    @staticmethod
    def _span_overlaps(spans: list[tuple[int, int]], start: int, length: int) -> bool:
        end = start + length
        for span_start, span_length in spans:
            span_end = span_start + span_length
            if start < span_end and end > span_start:
                return True
        return False


class MHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for Power Query M language."""

    def __init__(self, document):
        super().__init__(document)

        self.f_keyword = QTextCharFormat()
        self.f_keyword.setForeground(QColor('#C586C0'))
        self.f_keyword.setFontWeight(QFont.Weight.DemiBold)

        self.f_function = QTextCharFormat()
        self.f_function.setForeground(QColor('#4FC1FF'))

        self.f_string = QTextCharFormat()
        self.f_string.setForeground(QColor('#CE9178'))

        self.f_number = QTextCharFormat()
        self.f_number.setForeground(QColor('#B5CEA8'))

        self.f_comment = QTextCharFormat()
        self.f_comment.setForeground(QColor('#6A9955'))

        if M_KEYWORDS:
            kw_pattern = r"(?i)\b(" + "|".join(re.escape(word) for word in M_KEYWORDS) + r")\b"
            self.re_keyword = QRegularExpression(kw_pattern)
        else:
            self.re_keyword = None

        if M_FUNCTIONS:
            fn_pattern = r"(?i)\b(" + "|".join(re.escape(name) for name in M_FUNCTIONS) + r")\b(?=\s*\()"
            self.re_function = QRegularExpression(fn_pattern)
        else:
            self.re_function = None

        self.re_string = QRegularExpression(r'"[^"\\]*(?:\\.[^"\\]*)*"')
        self.re_number = QRegularExpression(r"\b[0-9]+(\.[0-9]+)?\b")
        self.re_line_comment = QRegularExpression(r"//.*$")
        self.re_comment_start = QRegularExpression(r"/\*")
        self.re_comment_end = QRegularExpression(r"\*/")

    def highlightBlock(self, text: str) -> None:
        self.setCurrentBlockState(0)

        comment_spans: list[tuple[int, int]] = []

        if self.previousBlockState() != 1:
            match = self.re_comment_start.match(text)
            start_index = match.capturedStart() if match.hasMatch() else -1
        else:
            start_index = 0

        while start_index >= 0:
            end_match = self.re_comment_end.match(text, start_index)
            end_index = end_match.capturedEnd() if end_match.hasMatch() else -1
            if end_index == -1:
                self.setCurrentBlockState(1)
                length = len(text) - start_index
                comment_spans.append((start_index, length))
                break
            else:
                length = end_index - start_index
                comment_spans.append((start_index, length))
                next_match = self.re_comment_start.match(text, end_index)
                start_index = next_match.capturedStart() if next_match.hasMatch() else -1

        if self.previousBlockState() == 1 and self.currentBlockState() != 1:
            self.setCurrentBlockState(0)

        it = self.re_line_comment.globalMatch(text)
        while it.hasNext():
            m = it.next()
            comment_spans.append((m.capturedStart(), m.capturedLength()))

        string_spans: list[tuple[int, int]] = []

        it = self.re_string.globalMatch(text)
        while it.hasNext():
            m = it.next()
            start = m.capturedStart()
            length = m.capturedLength()
            if self._span_overlaps(comment_spans, start, length):
                continue
            string_spans.append((start, length))
            if start > 0 and text[start - 1] == '#':
                continue
            self.setFormat(start, length, self.f_string)

        it = self.re_number.globalMatch(text)
        while it.hasNext():
            m = it.next()
            if not self._span_overlaps(comment_spans, m.capturedStart(), m.capturedLength()):
                self.setFormat(m.capturedStart(), m.capturedLength(), self.f_number)

        excluded_spans = comment_spans + string_spans

        if self.re_function:
            it = self.re_function.globalMatch(text)
            while it.hasNext():
                m = it.next()
                if not self._span_overlaps(excluded_spans, m.capturedStart(1), m.capturedLength(1)):
                    self.setFormat(m.capturedStart(1), m.capturedLength(1), self.f_function)

        if self.re_keyword:
            it = self.re_keyword.globalMatch(text)
            while it.hasNext():
                m = it.next()
                if not self._span_overlaps(excluded_spans, m.capturedStart(1), m.capturedLength(1)):
                    self.setFormat(m.capturedStart(1), m.capturedLength(1), self.f_keyword)

        for start, length in comment_spans:
            if length > 0:
                self.setFormat(start, length, self.f_comment)

    @staticmethod
    def _span_overlaps(spans: list[tuple[int, int]], start: int, length: int) -> bool:
        end = start + length
        for span_start, span_length in spans:
            span_end = span_start + span_length
            if start < span_end and end > span_start:
                return True
        return False


class DAXAutoCompleter(QObject):
    """Attach a QCompleter to a QTextEdit for DAX suggestions."""

    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor
        self.completer = QCompleter(sorted(set(DAX_KEYWORDS + DAX_FUNCTIONS)))
        self.completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self.completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        try:
            # Improves sorting with case-insensitive models
            self.completer.setModelSorting(QCompleter.ModelSorting.CaseInsensitivelySortedModel)
        except Exception:
            pass
        self.completer.setWidget(self.editor)
        self.completer.activated[str].connect(self.insert_completion)

        # Install event filter to trigger/show the completer
        self.editor.installEventFilter(self)

        # Distinguish function names for optional () insertion
        self.functions_set = set(DAX_FUNCTIONS)

        # Track last prefix to avoid redundant refresh
        self._last_prefix = None

    # Event filter must be on an QObject, so forward to instance method
    def eventFilter(self, obj, event):
        if obj is self.editor:
            et = event.type()
            if et == QEvent.Type.KeyPress:
                key = event.key()
                modifiers = event.modifiers()
                ctrl = modifiers & Qt.KeyboardModifier.ControlModifier

                if ctrl and key == Qt.Key.Key_Space:
                    prefix = self.current_word()
                    if prefix:
                        self.show_completions(prefix)
                    else:
                        self.completer.complete(self.editor.cursorRect())
                    return True

                # Hide on commit/navigation keys
                if key in (Qt.Key.Key_Escape, Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Tab):
                    self.completer.popup().hide()
                    return False

                # Update suggestions as user types letters, underscore, or deletes
                ch = event.text()
                if (ch and (ch.isalnum() or ch == '_')) or key in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
                    prefix = self.current_word()
                    if prefix and len(prefix) >= 1:
                        if prefix != self._last_prefix:
                            self._last_prefix = prefix
                            self.show_completions(prefix)
                    else:
                        self._last_prefix = None
                        self.completer.popup().hide()
                else:
                    # Non-word character typed -> hide suggestions
                    self._last_prefix = None
                    self.completer.popup().hide()
            elif et == QEvent.Type.FocusOut:
                self.completer.popup().hide()
        return False

    def current_word(self) -> str:
        """Return the word (letters/digits/underscore) surrounding the cursor."""
        cursor: QTextCursor = self.editor.textCursor()
        block_text = cursor.block().text()
        pos = cursor.positionInBlock()
        # Scan left
        l = pos
        while l > 0 and (block_text[l-1].isalnum() or block_text[l-1] == '_'):
            l -= 1
        # Scan right
        r = pos
        while r < len(block_text) and (block_text[r].isalnum() or block_text[r] == '_'):
            r += 1
        return block_text[l:r]

    def show_completions(self, prefix: str):
        self.completer.setCompletionPrefix(prefix)
        popup = self.completer.popup()
        # Ensure width fits content
        try:
            width = popup.sizeHintForColumn(0) + popup.verticalScrollBar().sizeHint().width() + 8
            rect = self.editor.cursorRect()
            rect.setWidth(width)
            self.completer.complete(rect)
        except Exception:
            self.completer.complete(self.editor.cursorRect())

    def insert_completion(self, completion: str):
        cursor: QTextCursor = self.editor.textCursor()
        cursor.beginEditBlock()

        # Replace the word under cursor with the completion
        cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.removeSelectedText()
        cursor.insertText(completion)

        # If it's a function and the next char isn't '(', insert parentheses
        if completion.upper() in self.functions_set:
            next_char = self._next_char_after_cursor(cursor)
            if next_char != '(':
                cursor.insertText('()')
                cursor.movePosition(QTextCursor.MoveOperation.Left)

        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)

    def _next_char_after_cursor(self, cursor: QTextCursor) -> str:
        temp = QTextCursor(cursor)
        temp.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, 1)
        return temp.selectedText()


@dataclass(frozen=True)
class LanguageDefinition:
    """Container describing syntax support for a code editor language."""

    name: str
    keywords: Sequence[str]
    functions: Sequence[str]
    highlighter_cls: type[QSyntaxHighlighter] | None = None
    line_comment: str | None = None
    extra_terms_provider: Callable[[], Iterable[str]] | None = None

    def completions(self) -> list[str]:
        """Return deduplicated, sorted completion terms."""
        terms = set(self.keywords or ())
        terms.update(self.functions or ())
        if self.extra_terms_provider:
            try:
                extra = self.extra_terms_provider() or ()
                terms.update(str(term) for term in extra if term)
            except Exception:
                pass
        return sorted(terms)


_LANGUAGE_DEFINITIONS: dict[str, LanguageDefinition] = {
    'dax': LanguageDefinition(
        name='dax',
        keywords=DAX_KEYWORDS,
        functions=DAX_FUNCTIONS,
        highlighter_cls=DAXHighlighter,
        line_comment='//',
        extra_terms_provider=_dax_model_terms,
    ),
    'm': LanguageDefinition(
        name='m',
        keywords=M_KEYWORDS,
        functions=M_FUNCTIONS,
        highlighter_cls=MHighlighter,
        line_comment='//',
    ),
}


def set_dax_model_identifiers(
    tables: Sequence[str] | None = None,
    columns: Sequence[str] | None = None,
) -> None:
    """Update the cached list of DAX model tables and columns used for completions."""

    def _normalize(items: Sequence[str] | None) -> list[str]:
        seen = set()
        normalized: list[str] = []
        if not items:
            return normalized
        for item in items:
            if item is None:
                continue
            text = str(item).strip()
            if not text or text in seen:
                continue
            seen.add(text)
            normalized.append(text)
        normalized.sort(key=str.casefold)
        return normalized

    new_tables = _normalize(tables)
    new_columns = _normalize(columns)

    global DAX_MODEL_TABLES, DAX_MODEL_COLUMNS
    if new_tables == DAX_MODEL_TABLES and new_columns == DAX_MODEL_COLUMNS:
        return

    DAX_MODEL_TABLES = new_tables
    DAX_MODEL_COLUMNS = new_columns

    try:
        from Coding.code_editor import CodeEditor

        CodeEditor.refresh_language("dax")
    except Exception:
        pass


def get_dax_model_identifiers() -> tuple[list[str], list[str]]:
    """Return copies of the current DAX model tables and columns lists."""
    return list(DAX_MODEL_TABLES), list(DAX_MODEL_COLUMNS)


def get_language_definition(language: str | None) -> LanguageDefinition | None:
    """Return the language definition for the given key, if available."""
    if not language:
        return None
    return _LANGUAGE_DEFINITIONS.get(language.lower())
