/*
 * Apache License
 * Version 2.0, January 2004
 * http://www.apache.org/licenses/
 *
 * Copyright (c) 2021, Elex
 * All rights reserved.
 */

package com.elex_project.xelhua;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.jetbrains.annotations.NotNull;

/**
 * Font builder
 *
 * @author Elex
 */
public final class FontBuilder {
	private final Font font;

	/**
	 * Font builder
	 *
	 * @param workbook workbook
	 */
	public FontBuilder(@NotNull final Workbook workbook) {
		this.font = workbook.createFont();
	}

	/**
	 * Font builder
	 *
	 * @param workbook workbook
	 * @param index    index of a font in workbook
	 */
	public FontBuilder(@NotNull final Workbook workbook, final int index) {
		this.font = workbook.getFontAt(index);
	}

	/**
	 * Font builder
	 *
	 * @param font font
	 */
	public FontBuilder(@NotNull final Font font) {
		this.font = font;
	}

	/**
	 * font family name
	 *
	 * @param fontName font family name
	 * @return builder
	 */
	@NotNull
	public FontBuilder name(@NotNull final String fontName) {
		font.setFontName(fontName);
		return this;
	}

	/**
	 * color
	 *
	 * @param color color
	 * @return builder
	 */
	@NotNull
	public FontBuilder color(@NotNull final IndexedColors color) {
		font.setColor(color.getIndex());
		return this;
	}

	/**
	 * bold
	 *
	 * @param bold bold?
	 * @return builder
	 */
	@NotNull
	public FontBuilder bold(final boolean bold) {
		font.setBold(bold);
		return this;
	}

	/**
	 * bold
	 *
	 * @return builder
	 */
	@NotNull
	public FontBuilder bold() {
		return bold(true);
	}

	/**
	 * italic
	 *
	 * @param italic italic
	 * @return builder
	 */
	@NotNull
	public FontBuilder italic(final boolean italic) {
		font.setItalic(italic);
		return this;
	}

	/**
	 * italic
	 *
	 * @return builder
	 */
	@NotNull
	public FontBuilder italic() {
		return italic(true);
	}

	/**
	 * strikeout
	 *
	 * @param strikeout strikeout
	 * @return builder
	 */
	@NotNull
	public FontBuilder strikeout(final boolean strikeout) {
		font.setStrikeout(strikeout);
		return this;
	}

	/**
	 * strikeout
	 *
	 * @return builder
	 */
	@NotNull
	public FontBuilder strikeout() {
		return strikeout(true);
	}

	/**
	 * underline
	 *
	 * @param underline underline
	 * @return builder
	 */
	@NotNull
	public FontBuilder underline(final boolean underline) {
		if (underline) {
			font.setUnderline(Font.U_SINGLE);
		} else {
			font.setUnderline(Font.U_NONE);
		}
		return this;
	}

	/**
	 * underline
	 *
	 * @return builder
	 */
	@NotNull
	public FontBuilder underline() {
		return underline(true);
	}

	/**
	 * height
	 *
	 * @param point size in points
	 * @return builder
	 */
	@NotNull
	public FontBuilder height(final float point) {
		font.setFontHeight((short) (point * 20));
		return this;
	}

	/**
	 * finish building a font
	 *
	 * @return font
	 */
	@NotNull
	public Font get() {
		return font;
	}
}
