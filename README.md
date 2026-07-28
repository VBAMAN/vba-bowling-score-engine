# VBA Bowling Score Engine

A bowling score calculation engine for Microsoft Excel VBA.

This project calculates bowling scores from a compact score string and supports:

* Strike scoring
* Spare scoring
* Open-frame scoring
* Frame-by-frame scores
* Cumulative scores
* Worksheet-based score display

The engine uses a score string such as:

```vb
scoreStr = "9/X8/X9/X8/X5/X5-"
```

The score can then be calculated with:

```vb
Debug.Print CalculateBowlingScore(scoreStr)
```

Result:

```text
195
```

A sample Excel workbook is included in the `sample` folder.

## Features

* Calculates the total score for a complete 10-frame bowling game
* Supports strikes (`X`)
* Supports spares (`/`)
* Supports open frames
* Supports missed pins (`-`)
* Calculates frame-by-frame scores
* Calculates cumulative scores
* Displays roll symbols and scores on an Excel worksheet
* Includes a sample Excel workbook
* Includes calculation test data

### Supported Score Symbols

The engine uses a compact score string containing bowling roll symbols.

### Supported Symbols

| Symbol  | Meaning                     |
| ------- | --------------------------- |
| `X`     | Strike                      |
| `/`     | Spare                       |
| `-`     | Miss                        |
| `G`     | Gutter ball                 |
| `0`–`9` | Number of pins knocked down |

### Example

```text
9/X8/X9/X8/X5/X5-
```

The score string is read from left to right as a sequence of bowling rolls.

The engine determines the frame structure according to standard 10-frame bowling rules.

### Perfect Game

```text
XXXXXXXXXXXX
```

Result:

```text
300
```

A perfect game contains 12 strikes:

* 10 strikes for the 10 frames
* 2 bonus strikes in the 10th frame

### All Spares

```text
9/9/9/9/9/9/9/9/9/9/9
```

Result:

```text
190
```

### Open Frames

```text
8172635445362718909-
```

Result:

```text
90
```

### Gutter Ball

`G` represents a gutter ball and is calculated as zero pins.

Example:

```text
G9G9G9G9G9G9G9G9G9G9
```

Result:

```text
90
```

### Important

The score string must not contain spaces.

Correct:

```text
8172635445362718909-
```

Not supported:

```text
81 72 63 54 45 36 27 18 90 9-
```

The current parser reads the score string character by character and does not remove spaces automatically.

## Module Structure

The project is organized into the following VBA modules:

| Module                   | Description                                 |
| ------------------------ | ------------------------------------------- |
| `BOWL_00_Init`           | Constants and initial settings              |
| `BOWL_10_Parse`          | Parses the bowling score string             |
| `BOWL_20_Calc`           | Calculates bowling scores                   |
| `BOWL_60_Display_Scores` | Displays roll symbols and cumulative scores |
| `BOWL_62_Display_Frames` | Displays bowling frame numbers              |
| `BOWL_75_Layout`         | Checks the scoreboard layout                |
| `BOWL_80_Worksheet`      | Draws and clears the worksheet scoreboard   |
| `BOWL_95_Run`            | Runs the sample score display               |
| `BOWL_98_Test`           | Contains score calculation test data        |

### Core Modules

The main calculation engine is contained in:

```text
BOWL_10_Parse
BOWL_20_Calc
```

The remaining modules provide worksheet display, layout utilities, sample execution, and test support.

## Sample Workbook

A sample Excel workbook is included in the `sample` folder.

```text
sample/
└── vba-bowling-score-engine.xlsm
```

Open the workbook in Microsoft Excel and enable macros.

To run the sample:

1. Open `vba-bowling-score-engine.xlsm`
2. Enable VBA macros
3. Open the Visual Basic Editor
4. Select `BOWL_95_Run`
5. Run `RunStringScore`

The sample displays:

* Bowling frame numbers
* Roll symbols
* Cumulative scores

You can change the score string in `RunStringScore` and run the procedure again.

```vb
scoreStr = "9/X8/X9/X8/X5/X5-"
```

The score display will be updated using the specified score string.

## Test Data

The project includes score calculation tests in:

```text
BOWL_98_Test
```

Run the following procedure:

```vb
TestCalculateBowlingScore
```

The test module includes the following score patterns.

| Test         | Score String            | Expected Score |
| ------------ | ----------------------- | -------------: |
| Perfect Game | `XXXXXXXXXXXX`          |            300 |
| All Spares   | `9/9/9/9/9/9/9/9/9/9/9` |            190 |
| Open Frames  | `8172635445362718909-`  |             90 |
| Mixed Score  | `9/X8/X9/X8/X5/X5-`     |            195 |

The test results are displayed in the VBA Immediate Window.

### NG Input Example

The current parser does not remove spaces automatically.

```text
81 72 63 54 45 36 27 18 90 9-
```

The score string must be written without spaces:

```text
8172635445362718909-
```

This behavior is included in the test module as a reference for the current input format.

## Requirements

* Microsoft Excel
* Visual Basic for Applications (VBA)
* Macro-enabled Excel workbook (`.xlsm`)

The sample workbook is provided in `.xlsm` format.

To use the sample workbook:

1. Open the workbook in Microsoft Excel
2. Enable macros when prompted
3. Run the sample procedure from the Visual Basic Editor

The engine is designed for use with Microsoft Excel VBA.

## License

This project is licensed under the MIT License.

You are free to use, modify, and distribute this project, including for commercial  
purposes, subject to the terms of the MIT License.

See the [LICENSE](../LICENSE) file for details.
