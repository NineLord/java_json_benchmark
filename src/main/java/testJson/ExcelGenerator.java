package testJson;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import utils.MathDataCollector;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.Map.Entry;
import java.util.function.Consumer;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class ExcelGenerator implements Closeable {

    public static void main(String[] args) {
        try (ExcelGenerator generator = new ExcelGenerator(
                "/PATH/TO/report2.xlsx",
                "/PATH/TO/Input/hugeJson_numberOfLetters8_depth10_children5.json",
                10,
                8,
                10,
                5
        )) {
            Map<String, Long> database = new HashMap<>();
            database.put(Reporter.MeasurementType.GENERATE_JSON.name(), 10L);
            database.put(Reporter.MeasurementType.ITERATE_ITERATIVELY.name(), 20L);
            database.put(Reporter.MeasurementType.ITERATE_RECURSIVELY.name(), 30L);
            database.put(Reporter.MeasurementType.DESERIALIZE_JSON.name(), 40L);
            database.put(Reporter.MeasurementType.SERIALIZE_JSON.name(), 50L);
            List<Pair<Long, Long>> pcUsage = new LinkedList<>();
            pcUsage.add(Pair.of(0L, 2000L));
            pcUsage.add(Pair.of(25L, 3000L));
            pcUsage.add(Pair.of(50L, 4000L));
            generator.appendWorksheet("Test 1", database.entrySet().stream(), pcUsage);
            Map<String, Long> database2 = new HashMap<>();
            database2.put(Reporter.MeasurementType.GENERATE_JSON.name(), 50L);
            database2.put(Reporter.MeasurementType.ITERATE_ITERATIVELY.name(), 70L);
            database2.put(Reporter.MeasurementType.ITERATE_RECURSIVELY.name(), 130L);
            database2.put(Reporter.MeasurementType.DESERIALIZE_JSON.name(), 180L);
            database2.put(Reporter.MeasurementType.SERIALIZE_JSON.name(), 5L);
            List<Pair<Long, Long>> pcUsage2 = new LinkedList<>();
            pcUsage2.add(Pair.of(30L, 1000L));
            pcUsage2.add(Pair.of(15L, 2500L));
            pcUsage2.add(Pair.of(100L, 3000L));
            generator.appendWorksheet("Test 2", database2.entrySet().stream(), pcUsage2);
        } catch (IOException exception) {
            exception.printStackTrace();
        }
    }

    @NotNull private final String pathToSaveFile;
    @NotNull private final String jsonPath;
    private final int sampleInterval;
    private final int numberOfLetters;
    private final int depth;
    private final int numberOfChildren;
    @NotNull private final Workbook workbook;

    @NotNull private final MathDataCollector averageGeneratingJsons;
	@NotNull private final MathDataCollector averageIteratingJsonsIteratively;
	@NotNull private final MathDataCollector averageIteratingJsonsRecursively;
	@NotNull private final MathDataCollector averageDeserializingJsons;
	@NotNull private final MathDataCollector averageSerializingJsons;

	@NotNull private final MathDataCollector totalAverageCpu;
	@NotNull private final MathDataCollector totalAverageRam;

    @NotNull private final CellStyle styleBorder;
    @NotNull private final CellStyle styleBorderAndCenter;

    public ExcelGenerator(@NotNull final String pathToSaveFile, @NotNull final String jsonPath, final int sampleInterval, final int numberOfLetters, final int depth, final int numberOfChildren) {
        this.pathToSaveFile = pathToSaveFile;
        this.jsonPath = jsonPath;
        this.sampleInterval = sampleInterval;
        this.numberOfLetters = numberOfLetters;
        this.depth = depth;
        this.numberOfChildren = numberOfChildren;
        this.workbook = new XSSFWorkbook();

        this.averageGeneratingJsons = new MathDataCollector();
	    this.averageIteratingJsonsIteratively = new MathDataCollector();
	    this.averageIteratingJsonsRecursively = new MathDataCollector();
	    this.averageDeserializingJsons = new MathDataCollector();
	    this.averageSerializingJsons = new MathDataCollector();

	    this.totalAverageCpu = new MathDataCollector();
	    this.totalAverageRam = new MathDataCollector();

        this.styleBorder = this.workbook.createCellStyle();
        this.styleBorder.setBorderTop(BorderStyle.THIN);
        this.styleBorder.setBorderBottom(BorderStyle.THIN);
        this.styleBorder.setBorderLeft(BorderStyle.THIN);
        this.styleBorder.setBorderRight(BorderStyle.THIN);

        this.styleBorderAndCenter = this.workbook.createCellStyle();
        this.styleBorderAndCenter.setBorderTop(BorderStyle.THIN);
        this.styleBorderAndCenter.setBorderBottom(BorderStyle.THIN);
        this.styleBorderAndCenter.setBorderLeft(BorderStyle.THIN);
        this.styleBorderAndCenter.setBorderRight(BorderStyle.THIN);
        this.styleBorderAndCenter.setAlignment(HorizontalAlignment.CENTER);
        this.styleBorderAndCenter.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    //#region Adding Data
    public @NotNull ExcelGenerator appendWorksheet(@NotNull final String worksheetName,
                                 @NotNull final Stream<Entry<String, Long>> database,
                                 @NotNull final List<Pair<Long, Long>> pcUsage) {
        @NotNull final Sheet worksheet = this.workbook.createSheet(worksheetName);
        worksheet.createFreezePane(0, 1);

        generateTitles(worksheet);
        @NotNull final Pair<MathDataCollector, MathDataCollector> dataCollectors = this.addData(worksheet, database, pcUsage);
        this.addStyleToCells(worksheet, dataCollectors);
        resizeColumns(worksheet);

        return this;
    }

    private static void generateTitles(@NotNull final Sheet worksheet) {
        IntStream.range(0, 10).forEach(worksheet::createRow);

        //#region Column 1
        //#region Table 1
        worksheet.getRow(0).createCell(0).setCellValue("Title");
        worksheet.getRow(1).createCell(0).setCellValue("Generating JSON");
        worksheet.getRow(2).createCell(0).setCellValue("Iterating JSON Iteratively - BFS");
        worksheet.getRow(3).createCell(0).setCellValue("Iterating JSON Recursively - DFS");
        worksheet.getRow(4).createCell(0).setCellValue("Deserializing JSON");
        worksheet.getRow(5).createCell(0).setCellValue("Serializing JSON");
        worksheet.getRow(6).createCell(0).setCellValue("Total");
        //#endregion

        //#region Table 2
        worksheet.getRow(8).createCell(0).setCellValue("Average CPU (%)");
        worksheet.getRow(9).createCell(0).setCellValue("Average RAM (MB)");
        //#endregion
        //#endregion

        //#region Column 2
        worksheet.getRow(0).createCell(1).setCellValue("Time (ms)");
        //#endregion

        //#region Column 4
        worksheet.getRow(0).createCell(3).setCellValue("CPU (%)");
        //#endregion

        //#region Column 5
        worksheet.getRow(0).createCell(4).setCellValue("RAM (MB)");
        //#endregion
    }

    private @NotNull Pair<MathDataCollector, MathDataCollector> addData(@NotNull final Sheet worksheet,
                                                            @NotNull final Stream<Entry<String, Long>> database,
                                                            @NotNull final List<Pair<Long, Long>> pcUsage) {
        @NotNull final MathDataCollector rowTotal = new MathDataCollector();
        @NotNull final MathDataCollector columnCpuUsage = new MathDataCollector();
        @NotNull final MathDataCollector columnRamUsage = new MathDataCollector();

        //#region JSON Manipulations
        database.forEach(entry -> {
            @NotNull final String testName = entry.getKey();
            final long testResult = entry.getValue();

            switch (Reporter.MeasurementType.valueOf(testName)) {
                case GENERATE_JSON:
                    worksheet.getRow(1).createCell(1).setCellValue(testResult);
                    this.averageGeneratingJsons.add(testResult);
                    break;
                case ITERATE_ITERATIVELY:
                    worksheet.getRow(2).createCell(1).setCellValue(testResult);
                    this.averageIteratingJsonsIteratively.add(testResult);
                    break;
                case ITERATE_RECURSIVELY:
                    worksheet.getRow(3).createCell(1).setCellValue(testResult);
                    this.averageIteratingJsonsRecursively.add(testResult);
                    break;
                case DESERIALIZE_JSON:
                    worksheet.getRow(4).createCell(1).setCellValue(testResult);
                    this.averageDeserializingJsons.add(testResult);
                    break;
                case SERIALIZE_JSON:
                    worksheet.getRow(5).createCell(1).setCellValue(testResult);
                    this.averageSerializingJsons.add(testResult);
                    break;
                default:
                    throw new RuntimeException("Invalid type of test: " + testName);
            }
            rowTotal.add(testResult);
        });

        worksheet.getRow(6).createCell(1).setCellValue(rowTotal.getSum());
        //#endregion

        //#region PC Usage
        int currentRowNumber = 1;
        for (@NotNull final Pair<Long, Long> pair : pcUsage) {
            final long cpu = pair.getLeft();
            final long ram = pair.getRight();

            @NotNull final Row currentRow = ExcelGenerator.getOrCreateRow(worksheet, currentRowNumber);
            currentRow.createCell(3).setCellValue(cpu);
            currentRow.createCell(4).setCellValue(ram);

            columnCpuUsage.add(cpu);
            columnRamUsage.add(ram);

            this.totalAverageCpu.add(cpu);
            this.totalAverageRam.add(ram);

            ++currentRowNumber;
        }

        Double columnCpuUsageAverage = columnCpuUsage.getAverage();
        if (columnCpuUsageAverage != null)
            worksheet.getRow(8).createCell(1).setCellValue(columnCpuUsageAverage);

        Double columnRamUsageAverage = columnRamUsage.getAverage();
        if (columnRamUsageAverage != null)
            worksheet.getRow(9).createCell(1).setCellValue(columnRamUsageAverage);
        //#endregion

        return Pair.of(columnCpuUsage, columnRamUsage);
    }

    private void addStyleToCells(@NotNull final Sheet worksheet,
                                 @NotNull final Pair<MathDataCollector, MathDataCollector> dataCollectors) {
        forEachCell(worksheet, 0, 0, 0, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));
        forEachCell(worksheet, 0, 3, 0, 4, cell -> cell.setCellStyle(this.styleBorderAndCenter));
        forEachCell(worksheet, 1, 0, 6, 0, cell -> cell.setCellStyle(this.styleBorder));
        forEachCell(worksheet, 1, 1, 6, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));
        forEachCell(worksheet, 8, 0, 9, 0, cell -> cell.setCellStyle(this.styleBorder));
        forEachCell(worksheet, 8, 1, 9, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));
        forEachCell(worksheet, 1, 3, worksheet.getPhysicalNumberOfRows(), 4, cell -> cell.setCellStyle(this.styleBorderAndCenter));

        SheetConditionalFormatting worksheetConditionalFormatting = worksheet.getSheetConditionalFormatting();

        Arrays.asList(
                Pair.of(dataCollectors.getLeft(), new CellRangeAddress[]{ new CellRangeAddress(1, worksheet.getPhysicalNumberOfRows(), 3, 3) }),
                Pair.of(dataCollectors.getRight(), new CellRangeAddress[]{ new CellRangeAddress(1, worksheet.getPhysicalNumberOfRows(), 4, 4) })
        ).forEach(pair -> {
            MathDataCollector mathDataCollector = pair.getLeft();
            CellRangeAddress[] rangeAddresses = pair.getRight();

            ConditionalFormattingRule rule = worksheetConditionalFormatting.createConditionalFormattingColorScaleRule();
            ColorScaleFormatting colorScaleFormatting = rule.getColorScaleFormatting();

            colorScaleFormatting.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            colorScaleFormatting.getThresholds()[0].setValue(mathDataCollector.getMinimum());
            colorScaleFormatting.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            colorScaleFormatting.getThresholds()[1].setValue(mathDataCollector.getAverage());
            colorScaleFormatting.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            colorScaleFormatting.getThresholds()[2].setValue(mathDataCollector.getMaximum());

            worksheetConditionalFormatting.addConditionalFormatting(rangeAddresses, rule);
        });
    }
    //#endregion

    //#region Add summary worksheet
    private void addAverageWorksheet() {
        @NotNull final Sheet worksheet = this.workbook.createSheet("Average");
        generateAverageTitles(worksheet);
        this.addAverageData(worksheet);
        addAverageStyleToCells(worksheet);
        resizeColumns(worksheet);
    }

    private static void generateAverageTitles(@NotNull final Sheet worksheet) {
        IntStream.range(0, 10).forEach(worksheet::createRow);

        //#region Column 1
        //#region Table 1
        worksheet.getRow(0).createCell(0).setCellValue("Title");
        worksheet.getRow(1).createCell(0).setCellValue("Average Generating JSONs");
        worksheet.getRow(2).createCell(0).setCellValue("Average Iterating JSONs Iteratively - BFS");
        worksheet.getRow(3).createCell(0).setCellValue("Average Iterating JSONs Recursively - DFS");
        worksheet.getRow(4).createCell(0).setCellValue("Average Deserializing JSONs");
        worksheet.getRow(5).createCell(0).setCellValue("Average Serializing JSONs");
        worksheet.getRow(6).createCell(0).setCellValue("Average Totals");
        //#endregion

        //#region Table 2
        worksheet.getRow(8).createCell(0).setCellValue("Average Total CPU (%)");
        worksheet.getRow(9).createCell(0).setCellValue("Average Total RAM (MB)");
        //#endregion
        //#endregion

        //#region Column 2
        worksheet.getRow(0).createCell(1).setCellValue("Time (ms)");
        //#endregion
    }

    private void addAverageData(@NotNull final Sheet worksheet) {
        @NotNull final MathDataCollector totalAverages = new MathDataCollector();
        Map<Integer, Double> cells = new HashMap<>();
        cells.put(1, this.averageGeneratingJsons.getAverage());
        cells.put(2, this.averageIteratingJsonsIteratively.getAverage());
        cells.put(3, this.averageIteratingJsonsRecursively.getAverage());
        cells.put(4, this.averageDeserializingJsons.getAverage());
        cells.put(5, this.averageSerializingJsons.getAverage());
        cells.forEach((cellRow, average) -> {
            if (average != null) {
                worksheet.getRow(cellRow).createCell(1).setCellValue(average);
                totalAverages.add(average);
            }
        });
        worksheet.getRow(6).createCell(1).setCellValue(totalAverages.getSum());

        Double totalAverageCpu = this.totalAverageCpu.getAverage();
        if (totalAverageCpu != null)
            worksheet.getRow(8).createCell(1).setCellValue(totalAverageCpu);

        Double totalAverageRam = this.totalAverageRam.getAverage();
        if (totalAverageRam != null)
            worksheet.getRow(9).createCell(1).setCellValue(totalAverageRam);
    }

    private void addAverageStyleToCells(@NotNull final Sheet worksheet) {
        forEachCell(worksheet, 0, 0, 0, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));
        forEachCell(worksheet, 1, 0, 6, 0, cell -> cell.setCellStyle(this.styleBorder));
        forEachCell(worksheet, 1, 1, 6, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));
        forEachCell(worksheet, 8, 0, 9, 0, cell -> cell.setCellStyle(this.styleBorder));
        forEachCell(worksheet, 8, 1, 9, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));
    }
    //#endregion

    //#region Add about worksheet
    private void createAboutWorksheet() {
        @NotNull final Sheet worksheet = this.workbook.createSheet("About");

        getOrCreateCell(worksheet, 0, 0).setCellValue("Path to JSON to be tested on (Iterating/Deserializing/Serializing)");
        getOrCreateCell(worksheet, 0, 1).setCellValue(this.jsonPath);

        getOrCreateCell(worksheet, 1, 0).setCellValue("CPU/RAM Sampling Interval (milliseconds)");
        getOrCreateCell(worksheet, 1, 1).setCellValue(this.sampleInterval);

        getOrCreateCell(worksheet, 2, 0).setCellValue("Number of letters to generate for each node in the generated JSON tree");
        getOrCreateCell(worksheet, 2, 1).setCellValue(this.numberOfLetters);

        getOrCreateCell(worksheet, 3, 0).setCellValue("Depth of the generated JSON tree");
        getOrCreateCell(worksheet, 3, 1).setCellValue(this.depth);

        getOrCreateCell(worksheet, 4, 0).setCellValue("Number of children each node in the generated JSON tree going to have");
        getOrCreateCell(worksheet, 4, 1).setCellValue(this.numberOfChildren);

        forEachCell(worksheet, 0, 0, 4, 1, cell -> cell.setCellStyle(this.styleBorder));
        forEachCell(worksheet, 0, 1, 4, 1, cell -> cell.setCellStyle(this.styleBorderAndCenter));

        resizeColumns(worksheet);
    }
    //#endregion

    //#region Excel Utils
    private static @NotNull Cell getOrCreateCell(@NotNull final Sheet worksheet, final int rowNumber, final int columnNumber) {
        @NotNull final Row row = getOrCreateRow(worksheet, rowNumber);
        return getOrCreateCell(row, columnNumber);
    }

    private static @NotNull Cell getOrCreateCell(@NotNull final Row row, final int columnNumber) {
        Cell cell = row.getCell(columnNumber);
        if (cell != null)
            return cell;
        else
            return row.createCell(columnNumber);
    }

    private static @NotNull Row getOrCreateRow(@NotNull final Sheet worksheet, final int rowNumber) {
        Row row = worksheet.getRow(rowNumber);
        if (row != null)
            return row;
        else
            return worksheet.createRow(rowNumber);
    }

    private static void resizeColumns(@NotNull final Sheet worksheet) {
        if (worksheet.getPhysicalNumberOfRows() <= 0)
            return;

        worksheet
                .getRow(worksheet.getFirstRowNum())
                .forEach(cell ->
                        worksheet.autoSizeColumn(cell.getColumnIndex())
                );
    }

    private static void forEachCell(@NotNull final Sheet worksheet,
                                    int startRow, int startColumn,
                                    int finishRow, int finishColumn, Consumer<Cell> callback) {
        for (int currentRow = startRow; currentRow <= finishRow; ++currentRow) {
            for (int currentColumn = startColumn; currentColumn <= finishColumn; ++currentColumn) {
                @NotNull final Cell currentCell = getOrCreateCell(worksheet, currentRow, currentColumn);
                callback.accept(currentCell);
            }
        }
    }
    //#endregion

    @Override
    public void close() throws IOException {
        try {
            this.addAverageWorksheet();
            this.createAboutWorksheet();
        } catch (Exception exception) {
            System.err.println("Unexpected error at adding summary/about worksheet: " + exception);
        }
        try (@NotNull final OutputStream outputStream = Files.newOutputStream(Paths.get(this.pathToSaveFile))) {
            this.workbook.write(outputStream);
        } finally {
            this.workbook.close();
        }
    }
}
