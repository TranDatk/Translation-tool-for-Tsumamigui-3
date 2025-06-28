# Tsumamigui 3 Translation Tool

A comprehensive tool for translating visual novel scenario files, specifically designed for Tsumamigui 3. This tool provides a complete workflow from extracting dialogues to packaging them back into the game.

## 🚀 Features

- **Multi-Tab Interface**: File Processing, Insert Again, Alice Tool
- **Configurable Dialogue Rules**: Customize dialogue delimiter patterns
- **Auto-Save Configuration**: Remembers your settings and file paths
- **Character Mapping**: Vietnamese to Japanese character replacement
- **Progress Tracking**: Real-time progress for all operations
- **Multi-Language Support**: English and Vietnamese interface
- **Complete Workflow**: TXT → Excel → TXT → AIN

## 📁 File Structure

```
Tsumamigui3Tool/
├── Tsumamigui3Tool.exe    # Main application (89MB)
├── vn_config.json         # Auto-generated config file
└── README.md             # This file
```

**Note**: `alice-tool` folder is embedded in the executable.

## 🔄 Complete Workflow

### Step 1: Extract Dialogues (File Processing Tab)
1. **Input**: Original TXT scenario file from game
2. **Output**: Excel file for translation work
3. **Process**: Parse and extract dialogue segments

### Step 2: Translate
1. Open the Excel file
2. Fill in translations in the "Translate" column
3. Use special values:
   - **Empty**: Skip this dialogue (keep original)
   - **"null"**: Uncomment but leave empty
   - **Text**: Your translation

### Step 3: Insert Translations (Insert Again Tab)
1. **Input**: Excel file with translations + Original TXT file
2. **Output**: Modified TXT file with translations
3. **Process**: Apply translations back to scenario file

### Step 4: Package for Game (Alice Tool Tab)
1. **Input**: AIN file + Translated TXT file
2. **Output**: New AIN file for game
3. **Process**: Compile into game-ready format

## 📖 Detailed Instructions

### 🎯 Tab 1: File Processing

**Purpose**: Extract dialogues from game scenario files for translation.

1. **Choose TXT file**: Select the original scenario file (e.g., `scenario.txt`)
2. **Choose Excel output**: Set where to save the extraction (e.g., `dialogues.xlsx`)
3. **Configure Rules**: Add dialogue delimiter patterns:
   - `「` / `」` (Japanese quotes)
   - `『` / `』` (Double quotes)
   - `（` / `）` (Parentheses)
   - `` / `。` (Empty start, period end)
4. **Rule Priority**: Use ↑/↓ to arrange rules (top = highest priority)
5. **Click Convert**: Extract dialogues to Excel

**Example Rules Setup**:
```
Priority 1: 『 → 』 (Narrative quotes)
Priority 2: 「 → 」 (Character dialogue)
Priority 3: （ → ） (Thoughts/effects)
Priority 4:   → 。 (General sentences)
```

### 🎯 Tab 2: Insert Again

**Purpose**: Apply translations back to the original scenario file.

1. **Choose Excel file**: Select the file with completed translations
2. **Choose TXT file**: Select the original scenario file to modify
3. **Configure Settings**:
   - **Max characters**: Line length limit (default: 50)
   - **Virtual Characters**: Vietnamese accented characters
   - **Physical Characters**: Japanese replacement characters
4. **Click Insert**: Apply translations

**Translation Column Values**:
- **Empty cell**: Keep original Japanese (stays commented `;m[...]`)
- **"null"**: Uncomment but empty (`m[123] = ""`)
- **Actual text**: Uncomment and insert translation (`m[123] = "Your translation"`)

**Character Mapping Example**:
```
Vietnamese: áàảãạ éèẻẽẹ íìỉĩị óòỏõọ úùủũụ ýỳỷỹỵ đ
Japanese:   ｱｱｱｱｱ ｴｴｴｴｴ ｲｲｲｲｲ ｵｵｵｵｵ ｳｳｳｳｳ ｲｲｲｲｲ ﾄﾞ
```

### 🎯 Tab 3: Alice Tool

**Purpose**: Package translated scenario into game-ready format.

1. **Choose Ain file**: Select the game's script file (e.g., `Tsumamigui3.ain`)
2. **Choose TXT file**: Select the scenario file with applied translations
3. **Choose Output path**: Where to save the new AIN file
4. **Click Pack Ain File**: Compile for game

**Command Generated**:
```bash
alice.exe ain edit -t [translated.txt] -o [output.ain] [input.ain]
```

## ⚙️ Configuration

All settings are automatically saved to `vn_config.json`:

```json
{
  "txt_path": "path/to/scenario.txt",
  "out_path": "path/to/output.xlsx",
  "rules": [
    {"start": "『", "end": "』"},
    {"start": "「", "end": "」"}
  ],
  "insert_config": {
    "max_chars": 50,
    "vir_chars": "áàảãạ...",
    "phy_chars": "｡ュョ､･..."
  },
  "alice_config": {
    "ain_file_path": "path/to/game.ain",
    "txt_file_path": "path/to/translated.txt",
    "output_ain_path": "path/to/output.ain"
  }
}
```

## 🔧 Example Workflow

### Sample Files:
- **Input**: `scenario.txt` (game scenario)
- **Work**: `translation.xlsx` (for translation)
- **Modified**: `scenario_translated.txt` (with translations)
- **Output**: `game_translated.ain` (final game file)

### Process:
1. **Extract**: `scenario.txt` → `translation.xlsx`
2. **Translate**: Fill Excel file with translations
3. **Insert**: `translation.xlsx` + `scenario.txt` → `scenario_translated.txt`
4. **Package**: `game.ain` + `scenario_translated.txt` → `game_translated.ain`

## 📊 Excel File Format

The generated Excel file has 4 columns:

| Range     | Speaker | Dialogue                | Translate           |
|-----------|---------|-------------------------|---------------------|
| 1069      | ナレーター | 「結婚！？」                 "Marriage!?"         
| 1070-1072 | 明人      裏返り、震える声が部屋に...   | Voice trembling...  |
| 1073      | ナレーター | よく晴れた、とある冬の日         null                

**Column Descriptions**:
- **Range**: Line numbers (single or range)
- **Speaker**: Character or narrator name
- **Dialogue**: Original Japanese text
- **Translate**: Your translation (fill this column)

## 🐛 Troubleshooting

### Common Issues:

**1. "alice.exe not found"**
- The executable should have alice-tool embedded
- If error persists, ensure you're using the full build

**2. "Cannot open TXT file"**
- Check file encoding (should be UTF-8)
- Ensure file is not locked by other applications

**3. "Excel file corrupted"**
- Re-extract from original TXT file
- Check if Excel file was saved properly

**4. "Translations not appearing in game"**
- Ensure AIN file is in correct game directory
- Backup original AIN file before replacing

**5. "Character encoding issues"**
- Check Virtual/Physical character mappings
- Ensure max characters setting is appropriate

### Performance Tips:

- **Large files**: Process in smaller chunks if needed
- **Memory usage**: Close other applications during processing
- **Speed**: Use SSD storage for better performance

## 📝 File Formats

### Supported Input:
- **TXT**: UTF-8 encoded scenario files
- **AIN**: Alice engine script files
- **XLSX**: Excel workbook files

### Generated Output:
- **XLSX**: Excel files with dialogue data
- **TXT**: Modified scenario files
- **AIN**: Compiled game script files

## 🌐 Language Support

- **Interface**: English / Vietnamese
- **Content**: Japanese (original) → Any target language
- **Character Sets**: Unicode support for all languages

## 📞 Support

For issues or questions:
1. Check this README first
2. Verify your workflow matches the examples
3. Check file formats and encodings
4. Test with smaller files first

## 📜 License

This tool is provided as-is for translation purposes. 

**Alice Tools**: The embedded alice.exe is from the Alice Tools project (Read more: https://haniwa.technology/alice-tools/README-ain.html).
**OpenPyXL**: Used for Excel file processing.
**Python**: Runtime environment.

---

**Version**: 1.0  
**Last Updated**: June 28, 2025  
**Compatibility**: Windows 10/11  

Made with ❤️ for the visual novel translation community. 