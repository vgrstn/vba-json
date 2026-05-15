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
- **Configurable** — seven `#Const` compiler directives control UTC conversion, date parsing, Collection keys, Excel large numbers, quoted/unquoted keys
- **High-performance string buffer** — `StringBuilder` UDT with exponential growth prevents O(n²) string concatenation
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
| `Collection` | json-object (set keys unique; unset keys become `""`) |
| Array | json-array (any number of dimensions) |
| `Date` (date + time, ≥ 1) | `"yyyy-mm-ddThh:mm:ssZ"` UTC string *(if `#Const jsonConvertUTC = True`)* |
| `Date` (date only, or < 1) | VB default string representation |
| `String` | json-string (Unicode/special chars escaped; `/` is **not** escaped — RFC 4627 §2.5 permits but does not require it) |
| `Long`, `Integer`, `Byte`, `LongLong` | json-number |
| `Single`, `Double`, `Currency`, `Decimal` | json-number |
| `Boolean` | `true` / `false` |
| `Null`, `Empty` | `null` |

### Deserialization (JSON → VB)

| JSON | VB type |
|---|---|
| json-object (unique non-empty keys) | `Dictionary` |
| json-object (empty or unique set keys) | `Collection` |
| json-array | `Variant()` array (n-dimensional) |
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
| `#Const API` | `False` | Use Windows API (`SysReAllocString`) or faster Variant ByRef approach to read Collection keys |
| `#Const jsonConvertUTC` | `True` | Convert `VB Date` (date + time) to/from UTC ISO-8601 string |
| `#Const jsonConvertDate` | `False` | Convert any json-string recognised by `VBA.IsDate` to a `VB Date` on parsing |
| `#Const jsonConvertCollectionKeys` | `True` | Include Collection keys in the serialized json-object |
| `#Const jsonConvertExcelLargeNumber` | `False` | Treat strings/numbers with more than 15 digits as Excel large numbers |
| `#Const jsonSerializeQuotedKeys` | `True` | Wrap object keys in quotes (standard JSON) |
| `#Const jsonDeserializeUnquotedKeys` | `False` | Also accept unquoted keys when parsing |

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
| VB Array | Stops after 2 dimensions | Any dimension |
| VB Collection | Serialized as json-array | Serialized as json-object |
| Collection keys | Not accessible | Serialized and deserialized |
| Dictionary keys | Not encoded | Encoded |
| `LongLong` support | No | Yes (x64) |
| json-array → VB | `Collection` | `Variant()` array |
| json-object (empty or unique set keys) | Stops at first empty key | Parsed to `Collection` |
| Minimal-type number | No | Yes |
| Large number | Parsed as `String` | Parsed as `Decimal` |
| Mac support | Yes | No (Windows only) |

---

## 🧠 Implementation notes

### Predeclared class
`Attribute VB_PredeclaredId = True` creates a module-level default instance; methods are called directly without `New`.

### StringBuilder
Serialization uses a pre-allocated string buffer (`StringBuilder` UDT) that doubles in size as needed, avoiding repeated string concatenation.

### Multi-dimensional arrays
Serialization uses `For Each` which enumerates in column-major order. Delimiters between elements (`,`, `],[`, `]],[[`, …) are derived from the array strides so no per-element branching is needed. Deserialization uses a 1-D buffer and a `SafeArrayCreate` pointer swap to redimension it in-place without copying the data.

### Collection keys
Collection keys are read from the Collection's internal doubly-linked list via `CopyMemory` + `SysReAllocString` (API mode) or a Variant ByRef memory construct (default, `#Const API = False`). The ByRef construct temporarily repoints a Variant's data address, letting VBA read any memory location without a COM-style BSTR allocation. This is approximately 5× faster than the API approach.

### Asymmetric `/` handling in Encode / Decode
RFC 4627 §2.5 permits but does not require escaping `/` as `\/`. `Encode` (serialization) does **not** emit `\/` — it is unnecessary and adds noise. `Decode` (deserialization) **does** accept `\/` — third-party producers (e.g. some .NET serializers) emit it, and the parser must remain interoperable. The lookup tables in the two functions differ by exactly this entry.

### UTC and DST
`DST()` calls `GetTimeZoneInformationForYear` (Vista+) to determine whether a given date falls in daylight saving time, using the machine's own timezone rules for any historical year. No hardcoded regional rules.

### `MatchValue` and structural characters
`MatchValue` scans for the first character in `"{}[],:"` to delimit a number or literal token. The double-quote `"` is intentionally absent: `MatchValue` is never called when the current character is `"` — that path dispatches to `MatchString` instead. Including `"` would cause spurious early termination on pathological inputs like an unquoted key immediately followed by a string value.

### `ArrDims`
Uses `LBound(arr, n)` in a loop with `On Error` to count dimensions. VBA raises on an out-of-range dimension index, making the error the natural loop-exit condition. Returns `0` for an empty (unallocated) array.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
