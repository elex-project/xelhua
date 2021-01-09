# Project Xelhua
Base Utility classes to manipulating Excel

## Gradle
```kotlin
repositories {
	mavenCentral()
	maven {
		name = "Github Packages"
		url = uri("https://maven.pkg.github.com/elex-project/xelhua")
		credentials {
			username = project.findProperty("github.username") as String
			password = project.findProperty("github.token") as String
		}
	}
}
dependencies {
	implementation("com.elex-project:xlsx:1.1.0")
	implementation("org.apache.poi:poi:4.1.2")
	implementation("org.apache.poi:poi-ooxml:4.1.2")
}
```

## Usage
```java
// 워크북 생성
Workbook workbook = createWorkbook();

// 시트 생성
Sheet sheet = getSheet(workbook, "Test 1");

// 행 가져오기
Row row = getRow(sheet, 0);

// 셀(0, 0) 가져오기, 셀에 문자열 쓰기
Cell cell = getCell(row, 0);
write(cell, "Hello");

// 셀(0, 1) 가져오기, 셀에 날짜 쓰기
cell = getCell(row, 1);
write(cell, LocalDate.now(), workbook);

// 셀(3, 3) 가져오기, 셀에 숫자 쓰기
cell = getCell(sheet, 3, 3);
write(cell, 123.45);
Font font = new FontBuilder(workbook)
        .height((short) 16)
        .color(IndexedColors.RED)
        .bold()
        .get();
CellStyle cellStyle = new CellStyleBuilder(workbook)
        .align(HorizontalAlignment.CENTER)
        .background(IndexedColors.YELLOW)
        .font(font)
        .get();
// 셀에 스타일 적용
cell.setCellStyle(cellStyle);

// 셀의 높이 설정
setHeight(cell, 20);

// 셀 합치기
mergeCells(sheet, 1, 2, 1, 1);

// 셀의 너비 자동 조정
autoWidth(sheet);

// 워크북을 파일로 저장
try (FileOutputStream fileOutputStream = new FileOutputStream(outFile1)) {
    writeOut(workbook, fileOutputStream);
    workbook.close();
}
```
---
developed by Elex

https://www.elex-project.com/
