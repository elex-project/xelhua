/*
 * Apache License
 * Version 2.0, January 2004
 * http://www.apache.org/licenses/
 *
 * Copyright (c) 2021, Elex
 * All rights reserved.
 */

package com.elex_project.xelhua;

import org.apache.poi.ss.usermodel.*;
import org.jetbrains.annotations.NotNull;

/**
 * CellStyle builder
 *
 * @author Elex
 */
public final class CellStyleBuilder {
	private final CellStyle cellStyle;

	/**
	 * CellStyle builder
	 *
	 * @param workbook workbook
	 */
	public CellStyleBuilder(@NotNull final Workbook workbook) {
		this.cellStyle = workbook.createCellStyle();
	}

	/**
	 * CellStyle builder
	 * @param workbook workbook
	 * @param index index of a style in workbook
	 */
	public CellStyleBuilder(@NotNull final Workbook workbook, final int index) {
		this.cellStyle = workbook.getCellStyleAt(index);
	}

	/**
	 * CellStyle builder
	 * @param cellStyle cellStyle
	 */
	public CellStyleBuilder(@NotNull final CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	/**
	 * CellStyle builder
	 * @param cell cell
	 */
	public CellStyleBuilder(@NotNull final Cell cell) {
		this.cellStyle = cell.getCellStyle();
	}

	/**
	 * background
	 *
	 * @param color       color
	 * @param fillPattern pattern
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder background(@NotNull final IndexedColors color, @NotNull final FillPatternType fillPattern) {
		cellStyle.setFillForegroundColor(color.getIndex());
		cellStyle.setFillPattern(fillPattern);
		return this;
	}

	/**
	 * background
	 *
	 * @param foregroundColor foreground Color
	 * @param backgroundColor background Color
	 * @param fillPattern     pattern
	 * @return builder
	 * @see CellStyle#setFillForegroundColor(short)
	 * @see CellStyle#setFillBackgroundColor(short)
	 * @see CellStyle#setFillPattern(FillPatternType)
	 */
	@NotNull
	public CellStyleBuilder background(@NotNull final IndexedColors foregroundColor, @NotNull final IndexedColors backgroundColor, @NotNull final FillPatternType fillPattern) {
		cellStyle.setFillForegroundColor(foregroundColor.getIndex());
		cellStyle.setFillBackgroundColor(backgroundColor.getIndex());
		cellStyle.setFillPattern(fillPattern);
		return this;
	}

	/**
	 * background with a solid pattern
	 *
	 * @param color color
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder background(@NotNull final IndexedColors color) {
		return background(color, FillPatternType.SOLID_FOREGROUND);
	}

	/**
	 * vertical alignment
	 *
	 * @param alignment alignment
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder align(@NotNull final VerticalAlignment alignment) {
		cellStyle.setVerticalAlignment(alignment);
		return this;
	}

	/**
	 * horizontal alignment
	 *
	 * @param alignment alignment
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder align(@NotNull final HorizontalAlignment alignment) {
		cellStyle.setAlignment(alignment);
		return this;
	}

	/**
	 * border
	 *
	 * @param borderStyle border
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderTop(@NotNull final BorderStyle borderStyle) {
		cellStyle.setBorderTop(borderStyle);
		return this;
	}

	/**
	 * border
	 *
	 * @param color color
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderTop(@NotNull final IndexedColors color) {
		cellStyle.setTopBorderColor(color.getIndex());
		return this;
	}

	/**
	 * border
	 *
	 * @param borderStyle border
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderLeft(@NotNull final BorderStyle borderStyle) {
		cellStyle.setBorderLeft(borderStyle);
		return this;
	}

	/**
	 * border
	 *
	 * @param color color
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderLeft(@NotNull final IndexedColors color) {
		cellStyle.setLeftBorderColor(color.getIndex());
		return this;
	}

	/**
	 * border
	 *
	 * @param borderStyle border
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderRight(@NotNull final BorderStyle borderStyle) {
		cellStyle.setBorderRight(borderStyle);
		return this;
	}

	/**
	 * border
	 *
	 * @param color color
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderRight(@NotNull final IndexedColors color) {
		cellStyle.setRightBorderColor(color.getIndex());
		return this;
	}

	/**
	 * border
	 *
	 * @param borderStyle border
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderBottom(@NotNull final BorderStyle borderStyle) {
		cellStyle.setBorderBottom(borderStyle);
		return this;
	}

	/**
	 * border
	 *
	 * @param color color
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder borderBottom(@NotNull final IndexedColors color) {
		cellStyle.setBottomBorderColor(color.getIndex());
		return this;
	}

	/**
	 * font
	 * you may preper to use a {@link FontBuilder}
	 *
	 * @param font font
	 * @return builder
	 */
	@NotNull
	public CellStyleBuilder font(@NotNull final Font font) {
		cellStyle.setFont(font);
		return this;
	}

	/**
	 * finish building a cell style
	 *
	 * @return cell style
	 */
	@NotNull
	public CellStyle get() {
		return cellStyle;
	}
}
