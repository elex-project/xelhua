package com.elex_project.xelhua;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;

/**
 * Base Utility class to manipulating Excel
 */
public interface ExcelBase {

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

	@NotNull
	public static HSSFWorkbook getHSSFWorkbook(@NotNull final InputStream inputStream) throws IOException {
		return new HSSFWorkbook(inputStream);
	}

	@NotNull
	public static XSSFWorkbook getXSSFWorkbook(@NotNull final InputStream inputStream) throws IOException {
		return new XSSFWorkbook(inputStream);
	}

	@NotNull
	public static Sheet getSheet(@NotNull final Workbook workbook, @NotNull final String name) {
		final Sheet sheet = workbook.getSheet(name);
		if (null == sheet) {
			return workbook.createSheet(name);
		} else {
			return sheet;
		}
	}

	@Nullable
	public static Sheet getSheetOrNull(@NotNull final Workbook workbook, @NotNull final String name) {
		return workbook.getSheet(name);
	}

	@NotNull
	public static Sheet getSheet(@NotNull final Workbook workbook, final int index) {
		try {
			return workbook.getSheetAt(index);
		} catch (IllegalArgumentException e) {
			return workbook.createSheet();
		}
	}

	@Nullable
	public static Sheet getSheetOrNull(@NotNull final Workbook workbook, final int index) {
		try {
			return workbook.getSheetAt(index);
		} catch (IllegalArgumentException e) {
			return null;
		}
	}

	@NotNull
	public static Row getRow(@NotNull final Sheet sheet, final int rowNum) {
		final Row row = sheet.getRow(rowNum);
		if (null == row) {
			return sheet.createRow(rowNum);
		} else {
			return row;
		}
	}

	@Nullable
	public static Row getRowOrNull(@NotNull final Sheet sheet, final int rowNum) {
		return sheet.getRow(rowNum);
	}

	@NotNull
	public static Cell getCell(@NotNull final Row row, final int colNum) {
		final Cell cell = row.getCell(colNum);
		if (null == cell) {
			return row.createCell(colNum);
		} else {
			return cell;
		}
	}

	@Nullable
	public static Cell getCellOrNull(@NotNull final Row row, final int colNum) {
		return row.getCell(colNum);
	}

	public static String readString(@NotNull final Cell cell) throws IllegalStateException {
		return cell.getStringCellValue();
	}

	public static double readNumeric(@NotNull final Cell cell) throws IllegalStateException, NumberFormatException {
		return cell.getNumericCellValue();
	}

	public static boolean readBoolean(@NotNull final Cell cell) throws IllegalStateException {
		return cell.getBooleanCellValue();
	}

	public static LocalDateTime readLocalDateTime(@NotNull final Cell cell) {
		if (DateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
		} else {
			throw new IllegalStateException();
		}
	}

	public static LocalDate readLocalDate(@NotNull final Cell cell) {
		if (DateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		} else {
			throw new IllegalStateException();
		}
	}

	@Nullable
	public static Comment readComment(@NotNull final Cell cell) {
		return cell.getCellComment();
	}

	public static void write(@NotNull Cell cell, final String value) {
		cell.setCellValue(value);
	}

	public static void write(@NotNull Cell cell, final double value) {
		cell.setCellValue(value);
	}

	public static void write(@NotNull Cell cell, final boolean value) {
		cell.setCellValue(value);
	}

	public static void write(@NotNull Cell cell, final LocalDate value, @NotNull final Workbook workbook, @NotNull final String format) {
		final CreationHelper creationHelper = workbook.getCreationHelper();
		final CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
		cell.setCellStyle(cellStyle);

		cell.setCellValue(value);
	}

	public static void write(@NotNull Cell cell, final LocalDate value, @NotNull final Workbook workbook) {
		write(cell, value, workbook, "yyyy-MM-dd");
	}

	public static void write(@NotNull Cell cell, final LocalDateTime value, @NotNull final Workbook workbook, @NotNull final String format) {
		final CreationHelper creationHelper = workbook.getCreationHelper();
		final CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
		cell.setCellStyle(cellStyle);

		cell.setCellValue(value);
	}

	public static void write(@NotNull Cell cell, final LocalDateTime value, @NotNull final Workbook workbook) {
		write(cell, value, workbook, "yyyy-MM-dd HH:mm:ss");
	}

	public static CellType getCellType(@NotNull final Cell cell) {
		return cell.getCellType();
	}

	public static CellStyle createCellStyle(@NotNull final Workbook workbook) {
		return workbook.createCellStyle();
	}

	public static Font createFont(@NotNull final Workbook workbook) {
		return workbook.createFont();
	}
}
