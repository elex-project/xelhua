/*
 * Apache License
 * Version 2.0, January 2004
 * http://www.apache.org/licenses/
 *
 * Copyright (c) 2021, Elex
 * All rights reserved.
 */

package com.elex_project.xelhua;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Iterator;

/**
 * Base Utility class to manipulating Excel
 *
 * @author Elex
 */
public interface ExcelBase {

	/**
	 * Open a Excel file
	 *
	 * @param file a file path with a extension .xls or .xlsx
	 * @return workbook
	 * @throws IOException if it can't be read from a file
	 */
	@NotNull
	public static Workbook getWorkbook(@NotNull final String file) throws IOException {
		final Workbook workbook;
		if (file.endsWith("xls")) {
			workbook = getHSSFWorkbook(new FileInputStream(file));
		} else {
			workbook = getXSSFWorkbook(new FileInputStream(file));
		}
		return workbook;
	}

	/**
	 * Create a new workbook
	 *
	 * @return xssf workbook
	 */
	public static Workbook createWorkbook() {
		return new XSSFWorkbook();
	}

	/**
	 * Open xls format from input stream
	 *
	 * @param inputStream maybe a file input stream
	 * @return workbook
	 * @throws IOException if it can not be read
	 */
	@NotNull
	public static HSSFWorkbook getHSSFWorkbook(@NotNull final InputStream inputStream) throws IOException {
		return new HSSFWorkbook(inputStream);
	}

	/**
	 * pen xlsx format from input stream
	 *
	 * @param inputStream maybe a file input stream
	 * @return workbook
	 * @throws IOException if it can not be read
	 */
	@NotNull
	public static XSSFWorkbook getXSSFWorkbook(@NotNull final InputStream inputStream) throws IOException {
		return new XSSFWorkbook(inputStream);
	}

	/**
	 * Get a named sheet from workbook, or create new one.
	 *
	 * @param workbook workbook
	 * @param name     name of a sheet
	 * @return sheet
	 */
	@NotNull
	public static Sheet getSheet(@NotNull final Workbook workbook, @NotNull final String name) {
		final Sheet sheet = workbook.getSheet(name);
		if (null == sheet) {
			return workbook.createSheet(name);
		} else {
			return sheet;
		}
	}

	/**
	 * Get a named sheet from workbook, or return null.
	 *
	 * @param workbook workbook
	 * @param name     name of sheet
	 * @return sheet or null
	 */
	@Nullable
	public static Sheet getSheetOrNull(@NotNull final Workbook workbook, @NotNull final String name) {
		return workbook.getSheet(name);
	}

	/**
	 * Get an n-th sheet from workbook, or create new one.
	 *
	 * @param workbook workbook
	 * @param index    index
	 * @return sheet
	 */
	@NotNull
	public static Sheet getSheet(@NotNull final Workbook workbook, final int index) {
		try {
			return workbook.getSheetAt(index);
		} catch (IllegalArgumentException e) {
			return workbook.createSheet();
		}
	}

	/**
	 * Get an n-th sheet from workbook, or return null.
	 *
	 * @param workbook workbook
	 * @param index    index
	 * @return sheet or null
	 */
	@Nullable
	public static Sheet getSheetOrNull(@NotNull final Workbook workbook, final int index) {
		try {
			return workbook.getSheetAt(index);
		} catch (IllegalArgumentException e) {
			return null;
		}
	}

	/**
	 * Create a new sheet
	 *
	 * @param workbook workbook
	 * @return sheet
	 */
	public static Sheet createSheet(@NotNull Workbook workbook) {
		return workbook.createSheet();
	}

	/**
	 * Get a row from sheet, or create new one.
	 *
	 * @param sheet  sheet
	 * @param rowNum row number
	 * @return row
	 */
	@NotNull
	public static Row getRow(@NotNull final Sheet sheet, final int rowNum) {
		final Row row = sheet.getRow(rowNum);
		if (null == row) {
			return sheet.createRow(rowNum);
		} else {
			return row;
		}
	}

	/**
	 * Get a row from sheet, or null.
	 *
	 * @param sheet  sheet
	 * @param rowNum row number
	 * @return row or null
	 */
	@Nullable
	public static Row getRowOrNull(@NotNull final Sheet sheet, final int rowNum) {
		return sheet.getRow(rowNum);
	}

	/**
	 * Get a cell from row, or create new one.
	 *
	 * @param row    row
	 * @param colNum column number
	 * @return cell
	 */
	@NotNull
	public static Cell getCell(@NotNull final Row row, final int colNum) {
		final Cell cell = row.getCell(colNum);
		if (null == cell) {
			return row.createCell(colNum);
		} else {
			return cell;
		}
	}

	/**
	 * Get a cell from row, or null.
	 *
	 * @param row    row
	 * @param colNum column number
	 * @return cell or null
	 */
	@Nullable
	public static Cell getCellOrNull(@NotNull final Row row, final int colNum) {
		return row.getCell(colNum);
	}

	/**
	 * Get a cell with a location, or create new one.
	 *
	 * @param sheet  sheet
	 * @param rowNum row number
	 * @param colNum column number
	 * @return cell
	 */
	public static Cell getCell(@NotNull Sheet sheet, final int rowNum, final int colNum) {
		return getCell(getRow(sheet, rowNum), colNum);
	}

	/**
	 * Get a cell with a header name
	 *
	 * @param row       row to get a cell
	 * @param name      column name in a header row
	 * @param headerRow header row with names
	 * @return a cell or create new one.
	 * @throws IllegalStateException Couldn't find a cell with that name in header row.
	 */
	@NotNull
	public static Cell getCell(@NotNull final Row row, @NotNull final String name, @NotNull Row headerRow)
			throws IllegalStateException {
		final Iterator<Cell> iterator = headerRow.cellIterator();
		while (iterator.hasNext()) {
			final Cell headerColumn = iterator.next();
			String columnName;
			try {
				switch (headerColumn.getCellType()) {
					case NUMERIC:
						columnName = String.valueOf(headerColumn.getNumericCellValue());
						break;
					case STRING:
					case FORMULA:
						columnName = headerColumn.getStringCellValue();
						break;
					case BOOLEAN:
						columnName = String.valueOf(headerColumn.getBooleanCellValue());
						break;
					case BLANK:
					case ERROR:
					case _NONE:
					default:
						columnName = "";
						break;
				}
				if (name.equals(columnName)) { // found a matching column
					return getCell(row, headerColumn.getColumnIndex());
				}
			} catch (Throwable ignore) {
			}
		}
		throw new IllegalStateException("Couldn't find a cell with that name in header row.");
	}

	/**
	 * Read string value from a cell
	 *
	 * @param cell cell
	 * @return string
	 * @throws IllegalStateException if cannot read a value as a string
	 * @see Cell#getStringCellValue()
	 */
	public static String readString(@NotNull final Cell cell) throws IllegalStateException {
		return cell.getStringCellValue();
	}

	/**
	 * Read numeric value as double from a cell
	 *
	 * @param cell cell
	 * @return double
	 * @throws IllegalStateException if can not read a value as a number
	 * @throws NumberFormatException if can not read a value as a number
	 * @see Cell#getNumericCellValue()
	 */
	public static double readNumeric(@NotNull final Cell cell) throws IllegalStateException, NumberFormatException {
		return cell.getNumericCellValue();
	}

	/**
	 * Read boolean value from a cell
	 *
	 * @param cell cell
	 * @return boolean
	 * @throws IllegalStateException if can not read a value as a boolean
	 * @see Cell#getBooleanCellValue()
	 */
	public static boolean readBoolean(@NotNull final Cell cell) throws IllegalStateException {
		return cell.getBooleanCellValue();
	}

	/**
	 * Read date time from a cell
	 *
	 * @param cell cell
	 * @return local date time at system default zone
	 */
	public static LocalDateTime readLocalDateTime(@NotNull final Cell cell) throws IllegalStateException {
		if (DateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
		} else {
			throw new IllegalStateException("Cell is not formatted as a date.");
		}
	}

	/**
	 * Read date from a cell
	 *
	 * @param cell cell
	 * @return local date at system default zone
	 */
	public static LocalDate readLocalDate(@NotNull final Cell cell) throws IllegalStateException {
		if (DateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		} else {
			throw new IllegalStateException("Cell is not formatted as a date.");
		}
	}

	/**
	 * Read a comment from a cell
	 *
	 * @param cell cell
	 * @return comment
	 */
	@Nullable
	public static Comment readComment(@NotNull final Cell cell) {
		return cell.getCellComment();
	}

	/**
	 * Write a value
	 *
	 * @param cell  cell
	 * @param value string value
	 */
	public static void write(@NotNull Cell cell, final String value) {
		cell.setCellValue(value);
	}

	/**
	 * Write a value
	 *
	 * @param cell  cell
	 * @param value numeric value
	 */
	public static void write(@NotNull Cell cell, final double value) {
		cell.setCellValue(value);
	}

	/**
	 * Write a value
	 *
	 * @param cell  cell
	 * @param value boolean value
	 */
	public static void write(@NotNull Cell cell, final boolean value) {
		cell.setCellValue(value);
	}

	/**
	 * Write a value
	 *
	 * @param cell     cell
	 * @param value    date
	 * @param workbook workbook. it's required to generate a cell style.
	 * @param format   date pattern
	 */
	public static void write(@NotNull Cell cell, final LocalDate value, @NotNull final Workbook workbook, @NotNull final String format) {
		final CreationHelper creationHelper = workbook.getCreationHelper();
		final CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
		cell.setCellStyle(cellStyle);

		cell.setCellValue(value);
	}

	/**
	 * Write a value, with "yyyy-MM-dd" format.
	 *
	 * @param cell     cell
	 * @param value    date
	 * @param workbook workbook, it's required to generate a cell style.
	 */
	public static void write(@NotNull Cell cell, final LocalDate value, @NotNull final Workbook workbook) {
		write(cell, value, workbook, "yyyy-MM-dd");
	}

	/**
	 * Write a value
	 *
	 * @param cell     cell
	 * @param value    date time
	 * @param workbook workbook, it's required to generate a cell style.
	 * @param format   date time pattern
	 */
	public static void write(@NotNull Cell cell, final LocalDateTime value, @NotNull final Workbook workbook, @NotNull final String format) {
		final CreationHelper creationHelper = workbook.getCreationHelper();
		final CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
		cell.setCellStyle(cellStyle);

		cell.setCellValue(value);
	}

	/**
	 * Write a value, with "yyyy-MM-dd HH:mm:ss" format.
	 *
	 * @param cell     cell
	 * @param value    date time
	 * @param workbook workbook, it's required to generate a cell style.
	 */
	public static void write(@NotNull Cell cell, final LocalDateTime value, @NotNull final Workbook workbook) {
		write(cell, value, workbook, "yyyy-MM-dd HH:mm:ss");
	}

	/**
	 * Return a cell type. {@link Cell#getCellType()}
	 *
	 * @param cell cell
	 * @return cell type
	 */
	public static CellType getCellType(@NotNull final Cell cell) {
		return cell.getCellType();
	}

	/**
	 * Merge cells
	 *
	 * @param sheet    sheet
	 * @param firstRow first row number (inclusive)
	 * @param lastRow  last row number (inclusive)
	 * @param firstCol first column number (inclusive)
	 * @param lastCol  last column number (inclusive)
	 * @see Sheet#addMergedRegion(CellRangeAddress)
	 */
	public static void mergeCells(@NotNull final Sheet sheet,
	                              final int firstRow, final int lastRow, final int firstCol, final int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * set column width
	 *
	 * @param sheet sheet
	 * @param cell  cell
	 * @param chars number of charters
	 */
	public static void setWidth(@NotNull final Sheet sheet, @NotNull Cell cell, final int chars) {
		setWidth(sheet, cell.getColumnIndex(), chars);
	}

	/**
	 * set column width
	 *
	 * @param sheet  sheet
	 * @param colNum column number
	 * @param chars  number of charters
	 */
	public static void setWidth(@NotNull final Sheet sheet, final int colNum, final int chars) {
		sheet.setColumnWidth(colNum, (int) (chars * 256));
	}

	/**
	 * Set row height
	 *
	 * @param row    row
	 * @param points size in points
	 */
	public static void setHeight(@NotNull Row row, final float points) {
		row.setHeight((short) (points * 20));
	}

	/**
	 * Set row height
	 * @param cell cell
	 * @param points size in points
	 */
	public static void setHeight(@NotNull Cell cell, final float points) {
		setHeight(cell.getRow(), points);
	}

	/**
	 * Set row height
	 *
	 * @param sheet  sheet
	 * @param rowNum row number
	 * @param points size in points
	 */
	public static void setHeight(@NotNull final Sheet sheet, final int rowNum, final float points) {
		setHeight(getRow(sheet, rowNum), points);
	}

	/**
	 * set default width of cells in a sheet
	 *
	 * @param sheet sheet
	 * @param chars number of characters
	 */
	public static void setDefaultWidth(@NotNull final Sheet sheet, final int chars) {
		sheet.setDefaultColumnWidth(chars);
	}

	/**
	 * set default height of rows in a sheet
	 *
	 * @param sheet  sheet
	 * @param points size in points
	 */
	public static void setDefaultHeight(@NotNull final Sheet sheet, final int points) {
		sheet.setDefaultRowHeight((short) (points * 20));
	}

	/**
	 * auto-resize column width
	 *
	 * @param sheet  sheet
	 * @param colNum column number
	 */
	public static void autoWidth(@NotNull final Sheet sheet, final int colNum) {
		sheet.autoSizeColumn(colNum);
	}

	/**
	 * auto-resize column width
	 *
	 * @param sheet sheet
	 * @param cell  cell
	 */
	public static void autoWidth(@NotNull final Sheet sheet, @NotNull final Cell cell) {
		autoWidth(sheet, cell.getColumnIndex());
	}

	/**
	 * auto-resize width of all columns
	 * (repeat {@link #autoWidth(Sheet, Cell)} for cells in the first row)
	 *
	 * @param sheet sheet
	 */
	public static void autoWidth(@NotNull final Sheet sheet) {
		final Row row = getRow(sheet, 0);
		for (final Cell cell : row) {
			autoWidth(sheet, cell);
		}
	}

	/**
	 * Save workbook to output stream
	 * after finished, don't forget closing the output stream and workbook.
	 *
	 * @param workbook     workbook
	 * @param outputStream output stream
	 * @throws IOException couldn't write to
	 */
	public static void writeOut(@NotNull final Workbook workbook, @NotNull final OutputStream outputStream)
			throws IOException {
		workbook.write(outputStream);
	}

	/**
	 * Save workbook to file
	 * after finished, don't forget closing workbook.
	 *
	 * @param workbook workbook
	 * @param file     file
	 * @throws IOException couldn't write to
	 */
	public static void writeOut(@NotNull final Workbook workbook, @NotNull final File file)
			throws IOException {
		try (FileOutputStream outputStream = new FileOutputStream(file)) {
			workbook.write(outputStream);
		}
	}

	/**
	 * Save workbook to file.
	 * it creates a parent directory, if needed.
	 * after finished, don't forget closing workbook.
	 *
	 * @param workbook workbook
	 * @param fileName proper file name extension could be appended if needed.
	 * @throws IOException couldn't write to
	 */
	public static void writeOut(@NotNull final Workbook workbook, @NotNull final String fileName)
			throws IOException {
		final File file;
		if (workbook instanceof XSSFWorkbook && !fileName.endsWith(".xlsx")) {
			file = new File(fileName + ".xlsx");
		} else if (workbook instanceof HSSFWorkbook && !fileName.endsWith(".xls")) {
			file = new File(fileName + ".xls");
		} else {
			file = new File(fileName);
		}
		file.getParentFile().mkdirs();

		try (FileOutputStream outputStream = new FileOutputStream(file)) {
			workbook.write(outputStream);
		}

	}
}
