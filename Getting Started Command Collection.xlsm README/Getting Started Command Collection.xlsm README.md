# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Getting Started Command Collection.xlsm\*\* contains definitions for:

[11 Robot Commands](#command-definitions)<BR>[8 Robot Texts](#text-definitions)<BR>[3 Robot Images](#image-definitions)<BR>

<BR>

## Available Robot Commands

[Demo_Commands](#demo_commands) | [Formatting](#formatting) | [Fun_and_Games](#fun_and_games) | [Text_Editing](#text_editing)

### Demo\_Commands

| Name | Description |
| --- | --- |
| [Say Hello](#say-hello) | Opens a message box with "Hello, (contents of the named range) \!" |
| [Wassup\!](#wassup) | Wraps the active cell contents with "Wassup, \_ \!" |

### Formatting

| Name | Description |
| --- | --- |
| [Auto\-Fit Column Width](#auto-fit-column-width) | Adjusts the widths of the columns based on selected cells' content. |
| [Auto\-Fit Row Height](#auto-fit-row-height) | Adjusts the heights of the rows based on selected cells' content. |
| [Make Cell Pretty](#make-cell-pretty) | Applies the pretty formatting to the active cell |
| [Make Selection Pretty](#make-selection-pretty) | Applies the pretty selection formatting to the active selection |
| [Make Tab Pretty](#make-tab-pretty) | Applies the pretty tab formatting to the active tab |

### Fun\_and\_Games

| Name | Description |
| --- | --- |
| [Robot Sticker](#robot-sticker) | Sticks a Robot Sticker on the active workbook |

### Text\_Editing

| Name | Description |
| --- | --- |
| [FirstName LastName](#firstname-lastname) | Takes the name in the active cell and switches it to FirstName LastName |
| [LastName, FirstName](#lastname-firstname) | Takes the name in the active cell and switches it to LastName, FirstName |
| [Toggle Case](#toggle-case) | Switches the contents of the active cell(s) from UPPER CASE to lower case, or lower case to Proper Case, or Proper Case to UPPER CASE |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [FirstName LastName.md](#firstname-lastnamemd) | FirstName LastName Command |
| [FirstNameFirst.lambda](#firstnamefirstlambda) | Definition of FirstNameFirst lambda function. |
| [LastName FirstName.md](#lastname-firstnamemd) | LastName, FirstName Command |
| [LastNameFirst.lambda](#lastnamefirstlambda) | Definition of LastNameFirst lambda function. |
| [Say Hello.md](#say-hellomd) | Say Hello Command |
| [Toggle Case.md](#toggle-casemd) | Toggle Case Command |
| [ToggleCase.lambda](#togglecaselambda) | Definition of the ToggleCase Lambda function |
| [Wassup\!.md](#wassupmd) | Wassup\! Command |

<BR>

## Available Robot Images

| Name | Description |
| --- | --- |
| [FirstNameLastName](#firstnamelastname) | |
| [LastNameFirstName](#lastnamefirstname) | |
| [SayHello](#sayhello) | |

<BR>

## Command Definitions

<BR>

### Auto\-Fit Column Width

*Adjusts the widths of the columns based on selected cells' content.*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Formatting`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.AutoFitColumnWidthSelection](./VBA/modCommands.bas#L4)()</code> |
| Launch Codes | <code>af</code> |

[^Top](#oa-robot-definitions)

<BR>

### Auto\-Fit Row Height

*Adjusts the heights of the rows based on selected cells' content.*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Formatting`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.AutoFitRowHeightSelection](./VBA/modCommands.bas#L14)()</code> |
| Launch Codes | <code>ar</code> |

[^Top](#oa-robot-definitions)

<BR>

### FirstName LastName

*Takes the name in the active cell and switches it to FirstName LastName*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Formula Command` `#Text_Editing`</sup>

[FirstName LastName Command](<.\Documentation\FirstName LastName.md>)

| Property | Value |
| --- | --- |
| Formula | <code>\=FirstNameFirst(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [FirstNameFirst.lambda](#firstnamefirstlambda) |
| User Context Filter | ExcelActiveCellIsNotEmpty |
| Documentation | [FirstName LastName.md](#firstname-lastnamemd) |
| Launch Codes | <code>fn</code> |

[^Top](#oa-robot-definitions)

<BR>

### LastName, FirstName

*Takes the name in the active cell and switches it to LastName, FirstName*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Formula Command` `#Text_Editing`</sup>

[LastName, FirstName Command](<.\Documentation\LastName FirstName.md>)

| Property | Value |
| --- | --- |
| Formula | <code>\=LastNameFirst(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [LastNameFirst.lambda](#lastnamefirstlambda) |
| User Context Filter | ExcelActiveCellIsNotEmpty |
| Documentation | [LastName FirstName.md](#lastname-firstnamemd) |
| Launch Codes | <code>ln</code> |

[^Top](#oa-robot-definitions)

<BR>

### Make Cell Pretty

*Applies the pretty formatting to the active cell*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Formatting`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.MakeCellPretty](./VBA/modCommands.bas#L50)()</code> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>mp</code> |

[^Top](#oa-robot-definitions)

<BR>

### Make Selection Pretty

*Applies the pretty selection formatting to the active selection*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Formatting`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.MakeSelectionPretty](./VBA/modCommands.bas#L71)()</code> |
| Launch Codes | <code>ms</code> |

[^Top](#oa-robot-definitions)

<BR>

### Make Tab Pretty

*Applies the pretty tab formatting to the active tab*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Formatting`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.MakeTabPretty](./VBA/modCommands.bas#L99)()</code> |
| Launch Codes | <code>mt</code> |

[^Top](#oa-robot-definitions)

<BR>

### Robot Sticker

*Sticks a Robot Sticker on the active workbook*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Fun_and_Games`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.RobotSticker](./VBA/modCommands.bas#L114)()</code> |
| Launch Codes | <code>rs</code> |

[^Top](#oa-robot-definitions)

<BR>

### Say Hello

*Opens a message box with "Hello, (contents of the named range) \!"*

<sup>`@Getting Started Command Collection.xlsm` `!VBA Macro Command` `#Demo_Commands`</sup>

[Say Hello Command](<.\Documentation\Say Hello.md>)

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCommands.SayHello](./VBA/modCommands.bas#L31)()</code> |
| Documentation | [Say Hello.md](#say-hellomd) |
| Launch Codes | <code>sh</code> |

[^Top](#oa-robot-definitions)

<BR>

### Toggle Case

*Switches the contents of the active cell(s) from UPPER CASE to lower case, or lower case to Proper Case, or Proper Case to UPPER CASE*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Formula Command` `#Text_Editing`</sup>

[Toggle Case Command](<.\Documentation\Toggle Case.md>)

| Property | Value |
| --- | --- |
| Formula | <code>\=ToggleCase(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [ToggleCase.lambda](#togglecaselambda) |
| User Context Filter | ExcelActiveCellIsNotEmpty |
| Documentation | [Toggle Case.md](#toggle-casemd) |
| Launch Codes | <code>tc</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wassup\!

*Wraps the active cell contents with "Wassup, \_ \!"*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Formula Command` `#Demo_Commands`</sup>

[Wassup! Command](<.\Documentation\Wassup!.md>)

| Property | Value |
| --- | --- |
| Formula | <code>\="Wassup, "& \[\[ActiveCell::Formula\]\] &"\!"</code> |
| User Context Filter | ExcelActiveCellIsNotEmpty |
| Documentation | [Wassup!.md](#wassupmd) |
| Launch Codes | <code>wu</code> |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### FirstName LastName.md

*FirstName LastName Command*

<sup>`@Getting Started Command Collection.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FirstName LastName.md](<./Text/FirstName LastName.md>) |
| Value | <code>\# \*\*First Name Last Name\*\* Command</code><br><code></code><br><code>Switches the cell contents from LastName, FirstName to FirstName LastName. For cells that don't contain a comma, no change is made to the cell contents. This command is visible when the active cell is non\-empty.</code><br><code></code><br><code>\!\[FirstName LastName\](oarobot:\/\/FirstNameLastNameImg)</code> |
| Content Type | Markdown |
| Markdown Id | <code>FirstNameLastName</code> |

[^Top](#oa-robot-definitions)

<BR>

### FirstNameFirst.lambda

*Definition of FirstNameFirst lambda function.*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FirstNameFirst.lambda](<./Text/FirstNameFirst.lambda.txt>) |
| Value | <code>FirstNameFirst \= LAMBDA(NameInCell, LET(</code><br><code> \\\\LambdaName, "FirstNameFirst",</code><br><code> IF(</code><br><code> ISNUMBER(SEARCH(",", NameInCell)),</code><br><code> TEXTAFTER(NameInCell, ", ") & " " & TEXTBEFORE(NameInCell, ", "),</code><br><code> NameInCell</code><br><code> )</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>FirstNameFirst</code> |
| Markdown Id | <code>FirstNameFirstlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### LastName FirstName.md

*LastName, FirstName Command*

<sup>`@Getting Started Command Collection.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [LastName FirstName.md](<./Text/LastName FirstName.md>) |
| Value | <code>\# \*\*Last Name, First Name\*\* Command</code><br><code></code><br><code> Switches the cell contents from FirstName LastName to LastName, FirstName. For cells that already contain a comma, no change is made to the cell contents. This command is visible when the active cell is non\-empty.</code><br><code></code><br><code>\!\[Last Name, First Name\](oarobot:\/\/LastNameFirstNameImg)</code> |
| Content Type | Markdown |
| Markdown Id | <code>LastNameFirstName</code> |

[^Top](#oa-robot-definitions)

<BR>

### LastNameFirst.lambda

*Definition of LastNameFirst lambda function.*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [LastNameFirst.lambda](<./Text/LastNameFirst.lambda.txt>) |
| Value | <code>LastNameFirst \= LAMBDA(NameInCell, LET(</code><br><code> \\\\LambdaName, "LastNameFirst",</code><br><code> IF(</code><br><code> ISNUMBER(SEARCH(",", NameInCell)),</code><br><code> NameInCell,</code><br><code> TEXTAFTER(NameInCell, " ", \-1) & ", " & TEXTBEFORE(NameInCell, " ", \-1)</code><br><code> )</code><br><code>));</code> |
| Content Type | ExcelFormula |
| Location | <code>LastNameFirst</code> |
| Markdown Id | <code>LastNameFirstlambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### Say Hello.md

*Say Hello Command*

<sup>`@Getting Started Command Collection.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Say Hello.md](<./Text/Say Hello.md>) |
| Value | <code>\# \*\*Say Hello\*\* Command</code><br><code></code><br><code>Populates a VBA Message Box with "Hello, \_\!" and the contents of an Excel Workbook\-scoped named range, \*NameForHello\* in the active workbook. </code><br><code></code><br><code>\!\[Say Hello\](oarobot:\/\/SayHelloImg)</code> |
| Content Type | Markdown |
| Markdown Id | <code>SayHello</code> |

[^Top](#oa-robot-definitions)

<BR>

### Toggle Case.md

*Toggle Case Command*

<sup>`@Getting Started Command Collection.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Toggle Case.md](<./Text/Toggle Case.md>) |
| Value | <code>\# \*\*Toggle Case\*\* Command</code><br><code></code><br><code>Description: Switches the contents of the active cells from UPPER CASE to lower case, or lower case to Proper Case, or Proper Case to UPPER CASE. The active cell(s) must be non\-empty to enable the command.</code> |
| Content Type | Markdown |
| Markdown Id | <code>togglecasemd</code> |

[^Top](#oa-robot-definitions)

<BR>

### ToggleCase.lambda

*Definition of the ToggleCase Lambda function*

<sup>`@Getting Started Command Collection.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ToggleCase.lambda](<./Text/ToggleCase.lambda.txt>) |
| Value | <code>ToggleCase \= LAMBDA(x, IF(ISBLANK(x), "", IF(EXACT(x, UPPER(x)), LOWER(x), IF(EXACT(x, LOWER(x)), PROPER(x), UPPER(x))) ));</code> |
| Content Type | ExcelFormula |
| Location | <code>ToggleCase</code> |
| Markdown Id | <code>ToggleCaselambda</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wassup\!.md

*Wassup\! Command*

<sup>`@Getting Started Command Collection.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Wassup!.md](<./Text/Wassup!.md>) |
| Value | <code>\# \*\*Wassup\!\*\* Command</code><br><code></code><br><code>Wraps the active cell contents with "Wassup, \_ \!"</code><br><code>Context: The active cell(s) must be non\-empty.</code> |
| Content Type | Markdown |
| Markdown Id | <code>Wassupmd</code> |

[^Top](#oa-robot-definitions)

<BR>

## Image Definitions

<BR>

### FirstNameLastName

<sup>`@Getting Started Command Collection.xlsm` `!Default Image` </sup>

| Property | Value |
| --- | --- |
| Value | ![OARobotImage](oarobot://FirstNameLastNameImg) |
| Markdown Id | <code>FirstNameLastNameImg</code> |

[^Top](#oa-robot-definitions)

<BR>

### LastNameFirstName

<sup>`@Getting Started Command Collection.xlsm` `!Default Image` </sup>

| Property | Value |
| --- | --- |
| Value | ![OARobotImage](oarobot://LastNameFirstNameImg) |
| Markdown Id | <code>LastNameFirstNameImg</code> |

[^Top](#oa-robot-definitions)

<BR>

### SayHello

<sup>`@Getting Started Command Collection.xlsm` `!Default Image` </sup>

| Property | Value |
| --- | --- |
| Value | ![OARobotImage](oarobot://SayHelloImg) |
| Markdown Id | <code>SayHelloImg</code> |

[^Top](#oa-robot-definitions)
