# vba-json
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA predeclared Class for JSON — converts any VB data structure to and from JSON, including multi-dimensional arrays, Collections with keys, and UTC date-time strings.

Fully implements the [JSON specification](https://www.json.org/json-en.html). Pure VBA, zero dependencies, x86/x64 compatible.

---

## 📦 Features

- **Predeclared** — accessible globally as `JSON.Stringify(...)` without `New` or `Set`
- **Multi-dimensional arrays** — serializes and deserializes arrays of any dimension (VBA stops at 2 with other converters)
- **Collection with keys** — serializes `VB Collection` as a json-object including its keys; deserializes back to a `Collection` when keys are empty or duplicate
- **Minimal-type number parsing** — json-numbers become `Long`, `LongLong` (x64), `Double`, or `Decimal`, whichever fits
- **UTC date-time** — `VB Date` values with both date and time convert to `"yyyy-mm-ddThh:mm:ssZ"` and back (configurable)
- **Prettify / Simplify** — format json for display or strip whitespace for transport
- **Configurable** — eight `#Const` compiler directives control UTC conversion, date parsing, Collection keys, Excel large numbers, quoted/unquoted keys and how invalid data is reported
- **Fast** — Stringify 2 to 50 times and Parse 1.2 to 10 times faster than JsonConverter (see Performance)
- **Whitespace anywhere** — the parser skips whitespace between tokens itself, so the positions in a parse error are those of the json as given; any json-value is allowed at the top level
- **Invalid data raises an error** — a value without a json form (an unsupported object or type, a matrix with an empty dimension) raises an error by default
- x86 / x64 compatible via `LongPtr` and `#If Win64`
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Description |
|---|---|
| `JSON.cls` | Source file with [Rubberduck](https://rubberduckvba.com/) annotations (`'@Description`, `'@PredeclaredId`, `'@IgnoreModule`) |
| `JSON_WithAttributes.cls` | Ready-to-import version with VB attributes baked in — no Rubberduck required |

Both files are identical in behaviour. Import `JSON_WithAttributes.cls` if you are not using Rubberduck.

> **Requires** a reference to `Microsoft Scripting Runtime` (for `Dictionary`).

---

## ⚙️ Public Interface

| Member | Description |
|---|---|
| `Stringify(var)` | Serializes a VB data structure (Array, Dictionary, Collection) to a JSON string |
| `Parse(json)` | Deserializes a JSON string to the corresponding VB data structure |
| `Prettify(json [, indentation [, space [, compact]]])` | Formats a JSON string with indentation; `compact = True` keeps the innermost array dimension on one line |
| `Simplify(json)` | Strips all whitespace outside strings |

All methods are available directly on the predeclared instance: `JSON.Stringify(...)`.

---

## 🚀 Quick Start

```vb
' Serialize a Dictionary to JSON
Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
d("name") = "Alice"
d("age") = 30
Debug.Print JSON.Stringify(d)       ' -> {"name":"Alice","age":30}

' Serialize a 2-D array
Dim a(1 To 2, 1 To 3) As Variant
a(1,1)=1: a(1,2)=2: a(1,3)=3
a(2,1)=4: a(2,2)=5: a(2,3)=6
Debug.Print JSON.Stringify(a)       ' -> [[1,4],[2,5],[3,6]]

' Serialize a Collection with keys
Dim c As New Collection
c.Add "hello", "greeting"
c.Add "world", "subject"
Debug.Print JSON.Stringify(c)       ' -> {"greeting":"hello","subject":"world"}

' Parse JSON to a Dictionary
Dim j As String: j = "{""x"":1,""y"":2}"
Dim r As Object: Set r = JSON.Parse(j)
Debug.Print r("x")                  ' -> 1

' Parse JSON to an array
Dim arr As Variant
arr = JSON.Parse("[1,2,3]")
Debug.Print arr(0)                  ' -> 1

' Prettify
Debug.Print JSON.Prettify(JSON.Stringify(d))
' {
'     "name": "Alice",
'     "age": 30
' }
```

---

## 🔄 Type mapping

### Serialization (VB → JSON)

| VB type | JSON |
|---|---|
| `Dictionary` | json-object (all keys set and unique) |
| `Collection` | json-object (unset keys become `""`) |
| Array | json-array (any number of dimensions) |
| `Date` (date + time, ≥ 1) | `"yyyy-mm-ddThh:mm:ssZ"` UTC string *(if `#Const jsonConvertUTC = True`)* |
| `Date` (date only, or < 1) | VB default string representation |
| `String` | json-string (Unicode/special chars escaped; `/` is **not** escaped — RFC 8259 permits but does not require it) |
| `Long`, `Integer`, `Byte`, `LongLong` | json-number |
| `Single`, `Double`, `Currency`, `Decimal` | json-number |
| `Boolean` | `true` / `false` |
| `Null`, `Empty` | `null` |
| Other objects and types, an array with an empty dimension | error (or the message in the json, see `jsonRaiseDataErrors`) |

### Deserialization (JSON → VB)

| JSON | VB type |
|---|---|
| json-object (unique non-empty keys) | `Dictionary` |
| json-object (unique keys or empty keys) | `Collection` |
| json-array | `Variant()` array; nested json-arrays become a matrix when rectangular (see below) |
| UTC string `"####-##-##T##:##:##Z"` | `Date` *(if `#Const jsonConvertUTC = True`)* |
| Other date string (recognized by `VBA.IsDate`) | `Date` *(only if `#Const jsonConvertDate = True`)* |
| json-string | `String` |
| integer json-number | `Long` → `LongLong` (x64) → `Decimal` |
| decimal json-number | `Double` → `Decimal` |
| `true` / `false` | `Boolean` |
| `null` | `Null` |

---

## ⚙️ Compiler directives

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Use the Windows API (`CopyMemory`, `SysReAllocString`) or the faster Variant ByRef construct for reading Collection keys and for the pointer copies of a parsed matrix |
| `#Const jsonConvertUTC` | `True` | Convert `VB Date` (date + time) to/from UTC ISO-8601 string |
| `#Const jsonConvertDate` | `False` | Convert any json-string recognised by `VBA.IsDate` to a `VB Date` on parsing |
| `#Const jsonConvertCollectionKeys` | `True` | Include Collection keys in the serialized json-object |
| `#Const jsonConvertExcelLargeNumber` | `False` | Meant for Excel, which stores a number as a `Double` with 15 significant digits and has no `LongLong`: a json-number with more than 15 digits becomes a string (also one that fits a `LongLong`), and thousands separators are removed from a large-number string when serialized |
| `#Const jsonSerializeQuotedKeys` | `True` | Wrap object keys in quotes (standard JSON). When `False`, a key that is not a simple name (letters, digits, `_`, `$`, not starting with a digit) is still quoted, as in JSON5, so the json can be read back |
| `#Const jsonDeserializeUnquotedKeys` | `False` | Also accept unquoted keys when parsing; an unquoted key must be a simple name |
| `#Const jsonRaiseDataErrors` | `True` | Stringify raises an error for data without a json form, or (`False`) puts the error message in the json in place of the value and prints it |

### `jsonConvertDate` × `jsonConvertUTC` — string deserialization behaviour

| `jsonConvertDate` | `jsonConvertUTC` | Behaviour |
|---|---|---|
| `False` | `False` | All json-strings are kept as a VB `String`. |
| `False` | `True` | Only an exact UTC format (`####-##-##T##:##:##Z`) is converted to a VB `Date`. All other json-strings are kept as a VB `String`. |
| `True` | `False` | Any json-string recognised by `VBA.IsDate` is converted to a VB `Date`. All others kept as a VB `String`. |
| `True` | `True` | Any json-string recognised by `VBA.IsDate` is converted to a VB `Date`. If not, an exact UTC format is converted to a VB `Date`. All others kept as a VB `String`. |

> **Why `jsonConvertDate` defaults to `False`:** `VBA.IsDate` is locale-dependent. A string like `"12/01/2025"` is silently converted to a `Date` on a US locale but kept as a `String` on a European locale. Keeping it `False` ensures consistent cross-locale behaviour and prevents accidental date conversion of strings like API tokens or version numbers that happen to match a date pattern.

---

## 🆚 Comparison with VBA-JSON (JsonConverter)

|  | `JsonConverter` | `JSON` (this) |
|---|---|---|
| VB Array | 1 or 2 dimensions; 3 and more give wrong output without an error | Any dimension |
| VB Collection | Serialized as json-array | Serialized as json-object |
| Collection keys | Not accessible | Serialized and deserialized |
| Dictionary keys | Not encoded | Encoded |
| `LongLong` support | No | Yes (x64) |
| json-array → VB | `Collection` | `Variant()` array |
| json-object (empty keys) | Stops at first empty key | Parsed to `Collection` |
| Minimal-type number | No | Yes |
| Large number | Parsed as `String` | Parsed as `Decimal` |
| Arrays with a negative bound, `LongLong` arrays | `[]` or an empty string, no error | Serialized |
| Mac support | Yes | No (Windows only) |

---

## 🧠 Implementation notes

### Predeclared class
`Attribute VB_PredeclaredId = True` creates a module-level default instance; methods are called directly without `New`.

### StringBuilder
Serialization uses a pre-allocated string buffer (`StringBuilder` UDT) that doubles in size as needed; structural characters are written in place with the `Mid$` statement. A version that collects the pieces in a String array and joins them with `Join` was measured and is 12 to 50% slower: every piece, `,` and `"` included, is a separate string.

### Encode and Decode
`Encode` first tests with `Like` whether a string needs encoding at all (about 0.06 µs); only then does it read the characters from a byte array and write the result to a byte array. `Decode` jumps from `\` to `\` with `InStr`, copies the runs in between at once and decodes the escapes with lookup tables. Both lookup tables are filled once in `Class_Initialize`.

### Parsing nested json-arrays
The outermost rectangular part of nested json-arrays becomes a matrix; anything irregular below it (a jagged or a mixed part such as `[[1,2],3]`) stays nested arrays, each treated the same way again. `[[1,2],[3,4]]` is a 2 × 2 matrix, `[[1,2],[3]]` an array of two arrays.

### Measured VBA pitfalls avoided
- `VBA.InStr` (qualified) copies the whole string on every call: 330 µs for 1 MB; the plain `InStr` takes no measurable time.
- `TypeName` of a `Scripting.Dictionary` takes about 500 µs; `TypeOf` takes no measurable time.
- A `New Dictionary` takes about 470 µs while no other Dictionary exists, and 5 µs otherwise; the class keeps one alive.
- A filled String local costs about 0.65 µs when the procedure exits; the recursive `Deserialize` uses a Variant.

### Multi-dimensional arrays
Serialization uses `For Each` which enumerates in column-major order. Delimiters between elements (`,`, `],[`, `]],[[`, …) are derived from the array strides so no per-element branching is needed. Deserialization uses a 1-D buffer and a `SafeArrayCreate` pointer swap to redimension it in-place without copying the data.

### Collection keys
Collection keys are read from the Collection's internal doubly-linked list via `CopyMemory` + `SysReAllocString` (API mode) or a Variant ByRef memory construct (default, `#Const API = False`). The ByRef construct temporarily repoints a Variant's data address, letting VBA read any memory location without a COM-style BSTR allocation. This is approximately 5× faster than the API approach.

### Asymmetric `/` handling in Encode / Decode
RFC 8259 permits but does not require escaping `/` as `\/`. `Encode` (serialization) does **not** emit `\/` — it is unnecessary and adds noise. `Decode` (deserialization) **does** accept `\/` — third-party producers (e.g. some .NET serializers) emit it, and the parser must remain interoperable. The lookup tables in the two functions differ by exactly this entry.

### UTC and DST
`DST()` calls `GetTimeZoneInformationForYear` (Vista+) to determine whether a given date falls in daylight saving time, using the machine's own timezone rules for any historical year. No hardcoded regional rules.

### `MatchValue` and structural characters
`MatchValue` scans for the first character in `"{}[],:"`, whitespace or the end of the json to delimit a number or literal token. The double-quote `"` is intentionally absent: `MatchValue` is never called when the current character is `"` — that path dispatches to `MatchString` instead. Including `"` would cause spurious early termination on pathological inputs like an unquoted key immediately followed by a string value.

### `ArrDims`
Uses `LBound(arr, n)` in a loop with `On Error` to count dimensions. VBA raises on an out-of-range dimension index, making the error the natural loop-exit condition. Returns `0` for an empty (unallocated) array.

---

## ⏱️ Performance

Time per element in µs at 10,000 elements (Office 64-bit, best of several runs; the benchmark is in `BenchmarkJSON`, not in this repository):

| Document | Stringify | ConvertToJson | Parse | ParseJson |
|---|---|---|---|---|
| Array of numbers | 0.25 | 1.28 | 4.6 | 11.4 |
| Array of strings | 0.50 | 8.4 | 2.7 | 10.9 |
| Dictionary | 0.76 | 1.9 | 3.6 | 37.7 |
| Records (Dictionary of 5 fields) | 11 | 558 | 36 | 80 |

JsonConverter's `ParseJson` returns Collections, not arrays, and its `ConvertToJson` writes 2-D arrays row-major and 3-D arrays wrongly, so matrices are not compared.

---

## 📝 Changes

**2026-11-09**
- Faster: `TypeOf` instead of `TypeName`, a `Like` fast path and byte arrays in `Encode`, lookup tables and `InStr` jumps in `Decode` and `MatchString`, plain `InStr` instead of `VBA.InStr`, no `Simplify` pass before parsing, early-bound bookkeeping for matrices and a Dictionary kept alive (a small matrix parsed 1.4 ms → 40 µs).
- Fixed: a string of escaped characters only (error 5); a number between a `Long` and a `LongLong` became a `Decimal`; an invalid json-object returned `Empty` instead of `Nothing`; a top-level number or literal was rejected; `[1 2]`, `[`, `{` and `[1,` were accepted; mixed arrays such as `[[1,2],3]` failed; invalid escapes (`\u12`, `\uZZZZ`, `\é`) raised the wrong error or were accepted; `Prettify` decoded strings (invalid output), wrote only CR and used only the first character of the indentation; unquoted keys could not be read back.
- New: `#Const jsonRaiseDataErrors`; Stringify raises an error for data without a json form.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
