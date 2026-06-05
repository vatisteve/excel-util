# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

`excel-util` is a lightweight Java 8 library that wraps Apache POI to read, create, and update Excel files. It is published to Maven Central under `io.github.vatisteve:excel-util`.

## Commands

Build and test (Maven):

```bash
mvn clean install          # compile, test, install to local repo
mvn test                   # run all tests
mvn test -Dtest=ExcelWriterTest                    # run a single test class
mvn test -Dtest=ExcelWriterBugTest#testStyleBugInAddCell   # run a single test method
mvn -DperformRelease=true deploy   # release profile (id "release"): signs (gpg), attaches sources/javadoc, publishes to Central
```

Tests use **JUnit 5 (Jupiter)** (`org.junit.jupiter.api.*`) via the `junit-jupiter` aggregator, run by `maven-surefire-plugin`. Note Jupiter's `assertEquals(expected, actual, message)` argument order (message last, unlike JUnit 4).

Writer tests verify behavior by writing to a `byte[]` and re-reading it with Apache POI **directly** through the shared `ExcelTestSupport` helper (`assertFirstSheet`/`readCell`), so a loader bug can't mask a writer bug. Test classes are organized one-concern-each: `ExcelWriterValueTypeTest`/`ExcelWriterTemporalTypeTest` (the value-handler dispatch matrix), `ExcelWriterStateTest` (cursor/positioning/errors), `ExcelWriterCellOperationTest` (the `CellAttribute` operation hook), `ExcelWriterConfigurationTest`, `ExcelLoaderApiTest` (full getter matrix + error/cast paths), plus `ExcelHelperTest`, `CellAttributeTest`, and `ElementNotFoundExceptionTest`.

Gotcha worth knowing: `SXSSFSheet.getRow(...)` only sees rows still in the streaming window, so `ExcelWriter.startAtRow(int)` can revisit rows written earlier in the same session but **not** pre-existing rows loaded from a template.

## Architecture

The library is entered through one factory and splits into two independent capabilities (read vs. write). Everything lives under `io.github.vatisteve.utils.excel`.

- **`ExcelUtilsFactory`** — the only public construction point. Static methods return `ExcelLoader` (read) and `ExcelWriter` (write) instances. Callers should not instantiate the `*Impl` classes directly.

- **`loader/`** — reading. `ExcelLoader` (interface) / `ExcelLoaderImpl` expose a large matrix of typed getters (`getString`, `getInteger`, `getLong`, `getValue<T>`) overloaded across three addressing styles: by sheet-index, by sheet-name, and against a "default sheet" (set via `setDefaultSheet`). Each accepts either `(column, row)` ints or a POI `CellAddress`. Numeric reads funnel through `castToNumber` (POI stores numbers as `double`, then transformed via a `Function<Double,T>`); string reads through `castToString`. Cell access and raw value extraction are delegated to `helper/ExcelHelper`.

- **`writer/`** — writing. `ExcelWriter` / `ExcelWriterImpl` use a **stateful cursor** model: the impl tracks `sheet`, `currentRow`, `nextRowIdx`, `nextColumnIdx`, and `cellIncrement`. You position with `startNewRow` / `startAtRow` / `startAtSheet`, then append cells left-to-right with `addCell`; each `addCell`/`autoIncrementCell` advances `nextColumnIdx`. Backed by `SXSSFWorkbook` (streaming) — when created from a template `InputStream` it wraps an `XSSFWorkbook`. Output via `build()` (byte[]) or `build(OutputStream)`.
  - **Value type dispatch**: `ExcelWriterImpl.initValueHandlers()` builds a `Map<Class<?>, BiConsumer<Object,Cell>>` keyed by exact runtime class. `detachAndSetCellValue` looks up the exact class first, then falls back to an `isInstance` scan (handles subclasses like `java.sql.Date`), and finally to `toString()`. Temporal types (`Instant`, `ZonedDateTime`, `OffsetDateTime`, `LocalTime`) and `BigDecimal`/`BigInteger` are converted before being set. When adding support for a new type, register it here.
  - **`ExcelWriterConfiguration`** — interface of `default` methods for global concerns: `sheetName`, `timeFormat`, `zoneId` (for temporal conversion), `cellStyle`, `rowHeight`, and `excelHeader`. `DefaultConfiguration` is the no-override baseline used by the no-arg factory method. `ExcelHeader` (nested, builder-based) defines an optional header row written automatically at construction via `initHeader()`.
  - **`CellAttribute`** (builder-based, immutable) + **`CellOperation`** (functional interface) — the extension path for per-cell behavior. `CellOperation.operate(sheet, cell)` runs custom POI logic (e.g. `sheet.addMergedRegion(...)`) before the value is set; see README for the merge example.

- **`common/` + `ElementNotFoundException`** — `ExcelElement` (SHEET/ROW/COLUMN/CELL) and `ElementIdentifier` (NAME/POSITION) enums are combined by the checked `ElementNotFoundException` to produce descriptive "no SHEET-POSITION with [..]" messages. Loader/writer catch POI's `IllegalArgumentException` on bad sheet/row lookups and rethrow as this exception.

## Conventions

- All public types carry full Javadoc; maintain it when changing signatures.
- Both `ExcelLoader` and `ExcelWriter` are `Closeable` — they own a POI `Workbook`. Use try-with-resources.
- Target is Java 8 (`maven.compiler.source/target=1.8`); do not use newer language features.
