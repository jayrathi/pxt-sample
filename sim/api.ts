/// <reference path="../libs/core/enums.d.ts"/>

namespace pxsim.worksheet {
    /**
     * This is hop
     */
    //% blockId="sampleHop" block="hop %hop on color %color=colorNumberPicker"
    //% hop.fieldEditor="gridpicker"
    export function hop(hop: Hop, color: number) {

    }

    //% blockId=sampleOnLand block="on land"
    //% optionalVariableArgs
    export function onLand(handler: (height: number, more: number, most: number) => void) {

    }
}

namespace pxsim.workbook {
    /**
     * This is hop
     */
    //% blockId="sampleHop" block="hop %hop on color %color=colorNumberPicker"
    //% hop.fieldEditor="gridpicker"
    export function hop(hop: Hop, color: number) {

    }

    //% blockId=sampleOnLand block="on land"
    //% optionalVariableArgs
    export function onLand(handler: (height: number, more: number, most: number) => void) {

    }

        /**
     * Specifies if the workbook is in AutoSave mode.
     */
    //% block="Workbook $this(Workbook) getAutoSave"
    //% group="Workbook"
    //% weight="100"
    export function getAutoSave(): boolean {
        return true;
    }

}

namespace pxsim.range {
    /**
     * Moves the sprite forward
     * @param steps number of steps to move, eg: 1
     */
    //% weight=90
    //% blockId=sampleForward block="forward %steps"
    export function forwardAsync(steps: number) {
        return board().sprite.forwardAsync(steps)
    }

    /**
     * Moves the sprite forward
     * @param direction the direction to turn, eg: Direction.Left
     * @param angle degrees to turn, eg:90
     */
    //% weight=85
    //% blockId=sampleTurn block="turn %direction|by %angle degrees"
    //% angle.min=-180 angle.max=180
    export function turnAsync(direction: Direction, angle: number) {
        let b = board();

        if (direction == Direction.Left)
            b.sprite.angle -= angle;
        else
            b.sprite.angle += angle;
        return Promise.delay(400)
    }

    /**
     * Triggers when the turtle bumps a wall
     * @param handler 
     */
    //% blockId=onBump block="on bump"
    export function onBump(handler: RefAction) {
        let b = board();

        b.bus.listen("Turtle", "Bump", handler);
    }
}

namespace pxsim.loops {

    /**
     * Repeats the code forever in the background. On each iteration, allows other code to run.
     * @param body the code to repeat
     */
    //% help=functions/forever weight=55 blockGap=8
    //% blockId=device_forever block="forever" 
    export function forever(body: RefAction): void {
        thread.forever(body)
    }

    /**
     * Pause for the specified time in milliseconds
     * @param ms how long to pause for, eg: 100, 200, 500, 1000, 2000
     */
    //% help=functions/pause weight=54
    //% block="pause (ms) %pause" blockId=device_pause
    export function pauseAsync(ms: number) {
        return Promise.delay(ms)
    }
}

function logMsg(m:string) { console.log(m) }

namespace pxsim.console {
    /**
     * Print out message
     */
    //% 
    export function log(msg:string) {
        logMsg("CONSOLE: " + msg)
        // why doesn't that work?
        board().writeSerial(msg + "\n")
    }
}

namespace pxsim {
    /**
     * A ghost on the screen.
     */
    //%
    export class Sprite {
        /**
         * The X-coordiante
         */
        //%
        public x = 100;
         /**
         * The Y-coordiante
         */
        //%
        public y = 100;
        public angle = 90;
        
        constructor() {
        }
        
        private foobar() {}

        /**
         * Move the thing forward
         */
        //%
        public forwardAsync(steps: number) {
            let deg = this.angle / 180 * Math.PI;
            this.x += Math.cos(deg) * steps * 10;
            this.y += Math.sin(deg) * steps * 10;
            board().updateView();

            if (this.x < 0 || this.y < 0)
                board().bus.queue("TURTLE", "BUMP");

            return Promise.delay(400)
        }
    }
}

namespace pxsim.excelScript {
    /**
     * Creates a new sprite
     */
    //% blockId="sampleCreate" block="createSprite"
    export function createSprite(): Sprite {
        return new Sprite();
    }

        /**
     * Represents the Excel application that manages the workbook.
     */
    export class Application {
        /**
         * Returns the Excel calculation engine version used for the last full recalculation.
         */
        //% block="Application $this(Application) getCalculationEngineVersion"
        //% group="Application"
        //% weight="10"
        getCalculationEngineVersion(): number { return 0;}

        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in `ExcelScript.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         */
        //% block="Application $this(Application) getCalculationMode"
        //% group="Application"
        //% weight="10"
        getCalculationMode(): CalculationMode {}

        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in `ExcelScript.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         */
        //% block="Application $this(Application) setCalculationMode $calculationMode"
        //% group="Application"
        //% weight="10"
        setCalculationMode(calculationMode: CalculationMode): void {}

        /**
         * Returns the calculation state of the application. See `ExcelScript.CalculationState` for details.
         */
        //% block="Application $this(Application) getCalculationState"
        //% group="Application"
        //% weight="10"
        getCalculationState(): CalculationState {}

        /**
         * Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.
         */
        //% block="Application $this(Application) getCultureInfo"
        //% group="Application"
        //% weight="10"
        getCultureInfo(): CultureInfo {}

        /**
         * Gets the string used as the decimal separator for numeric values. This is based on the local Excel settings.
         */
        //% block="Application $this(Application) getDecimalSeparator"
        //% group="Application"
        //% weight="10"
        getDecimalSeparator(): string {return "";}

        /**
         * Returns the iterative calculation settings.
         * In Excel on Windows and Mac, the settings will apply to the Excel Application.
         * In Excel on the web and other platforms, the settings will apply to the active workbook.
         */
        //% block="Application $this(Application) getIterativeCalculation"
        //% group="Application"
        //% weight="10"
        getIterativeCalculation(): IterativeCalculation {}

        /**
         * Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on the local Excel settings.
         */
        //% block="Application $this(Application) getThousandsSeparator"
        //% group="Application"
        //% weight="10"
        getThousandsSeparator(): string {return "";}

        /**
         * Specifies if the system separators of Excel are enabled.
         * System separators include the decimal separator and thousands separator.
         */
        //% block="Application $this(Application) getUseSystemSeparators"
        //% group="Application"
        //% weight="10"
        getUseSystemSeparators(): boolean {return true;}

        /**
         * Recalculate all currently opened workbooks in Excel.
         * @param calculationType Specifies the calculation type to use. See `ExcelScript.CalculationType` for details.
         */
        //% block="Application $this(Application) calculate $calculationType"
        //% group="Application"
        //% weight="10";
        calculate(calculationType: CalculationType): void {}
    }

    /**
     * Represents the iterative calculation settings.
     */
    export class IterativeCalculation {
        /**
         * True if Excel will use iteration to resolve circular references.
         */
        //% block="IterativeCalculation $this(IterativeCalculation) getEnabled"
        //% group="IterativeCalculation"
        //% weight="10"
        getEnabled(): boolean {return true;}

        /**
         * True if Excel will use iteration to resolve circular references.
         */
        //% block="IterativeCalculation $this(IterativeCalculation) setEnabled $enabled"
        //% group="IterativeCalculation"
        //% weight="10"
        setEnabled(enabled: boolean): void {}

        /**
         * Specifies the maximum amount of change between each iteration as Excel resolves circular references.
         */
        //% block="IterativeCalculation $this(IterativeCalculation) getMaxChange"
        //% group="IterativeCalculation"
        //% weight="10"
        getMaxChange(): number {return 0;}

        /**
         * Specifies the maximum amount of change between each iteration as Excel resolves circular references.
         */
        //% block="IterativeCalculation $this(IterativeCalculation) setMaxChange $maxChange"
        //% group="IterativeCalculation"
        //% weight="10"
        setMaxChange(maxChange: number): void {}

        /**
         * Specifies the maximum number of iterations that Excel can use to resolve a circular reference.
         */
        //% block="IterativeCalculation $this(IterativeCalculation) getMaxIteration"
        //% group="IterativeCalculation"
        //% weight="10"
        getMaxIteration(): number {return 0;}

        /**
         * Specifies the maximum number of iterations that Excel can use to resolve a circular reference.
         */
        //% block="IterativeCalculation $this(IterativeCalculation) setMaxIteration $maxIteration"
        //% group="IterativeCalculation"
        //% weight="10"
        setMaxIteration(maxIteration: number): void {}
    }

}

//
// Enum
//

/**
 * Enum representing all accepted conditions by which a date filter can be applied.
 * Used to configure the type of PivotFilter that is applied to the field.
 */
enum DateFilterCondition {
    /**
     * `DateFilterCondition` is unknown or unsupported.
     */
    unknown,

    /**
     * Equals comparator criterion.
     *
     * Required Criteria: {`comparator`}.
     * Optional Criteria: {`wholeDays`, `exclusive`}.
     */
    equals,

    /**
     * Date is before comparator date.
     *
     * Required Criteria: {`comparator`}.
     * Optional Criteria: {`wholeDays`}.
     */
    before,

    /**
     * Date is before or equal to comparator date.
     *
     * Required Criteria: {`comparator`}.
     * Optional Criteria: {`wholeDays`}.
     */
    beforeOrEqualTo,

    /**
     * Date is after comparator date.
     *
     * Required Criteria: {`comparator`}.
     * Optional Criteria: {`wholeDays`}.
     */
    after,

    /**
     * Date is after or equal to comparator date.
     *
     * Required Criteria: {`comparator`}.
     * Optional Criteria: {`wholeDays`}.
     */
    afterOrEqualTo,

    /**
     * Between `lowerBound` and `upperBound` dates.
     *
     * Required Criteria: {`lowerBound`, `upperBound`}.
     * Optional Criteria: {`wholeDays`, `exclusive`}.
     */
    between,

    /**
     * Date is tomorrow.
     */
    tomorrow,

    /**
     * Date is today.
     */
    today,

    /**
     * Date is yesterday.
     */
    yesterday,

    /**
     * Date is next week.
     */
    nextWeek,

    /**
     * Date is this week.
     */
    thisWeek,

    /**
     * Date is last week.
     */
    lastWeek,

    /**
     * Date is next month.
     */
    nextMonth,

    /**
     * Date is this month.
     */
    thisMonth,

    /**
     * Date is last month.
     */
    lastMonth,

    /**
     * Date is next quarter.
     */
    nextQuarter,

    /**
     * Date is this quarter.
     */
    thisQuarter,

    /**
     * Date is last quarter.
     */
    lastQuarter,

    /**
     * Date is next year.
     */
    nextYear,

    /**
     * Date is this year.
     */
    thisYear,

    /**
     * Date is last year.
     */
    lastYear,

    /**
     * Date is in the same year to date.
     */
    yearToDate,

    /**
     * Date is in Quarter 1.
     */
    allDatesInPeriodQuarter1,

    /**
     * Date is in Quarter 2.
     */
    allDatesInPeriodQuarter2,

    /**
     * Date is in Quarter 3.
     */
    allDatesInPeriodQuarter3,

    /**
     * Date is in Quarter 4.
     */
    allDatesInPeriodQuarter4,

    /**
     * Date is in January.
     */
    allDatesInPeriodJanuary,

    /**
     * Date is in February.
     */
    allDatesInPeriodFebruary,

    /**
     * Date is in March.
     */
    allDatesInPeriodMarch,

    /**
     * Date is in April.
     */
    allDatesInPeriodApril,

    /**
     * Date is in May.
     */
    allDatesInPeriodMay,

    /**
     * Date is in June.
     */
    allDatesInPeriodJune,

    /**
     * Date is in July.
     */
    allDatesInPeriodJuly,

    /**
     * Date is in August.
     */
    allDatesInPeriodAugust,

    /**
     * Date is in September.
     */
    allDatesInPeriodSeptember,

    /**
     * Date is in October.
     */
    allDatesInPeriodOctober,

    /**
     * Date is in November.
     */
    allDatesInPeriodNovember,

    /**
     * Date is in December.
     */
    allDatesInPeriodDecember,
}

/**
 * Enum representing all accepted conditions by which a label filter can be applied.
 * Used to configure the type of PivotFilter that is applied to the field.
 * `PivotFilter.criteria.exclusive` can be set to `true` to invert many of these conditions.
 */
enum LabelFilterCondition {
    /**
     * `LabelFilterCondition` is unknown or unsupported.
     */
    unknown,

    /**
     * Equals comparator criterion.
     *
     * Required Criteria: {`comparator`}.
     * Optional Criteria: {`exclusive`}.
     */
    equals,

    /**
     * Label begins with substring criterion.
     *
     * Required Criteria: {`substring`}.
     * Optional Criteria: {`exclusive`}.
     */
    beginsWith,

    /**
     * Label ends with substring criterion.
     *
     * Required Criteria: {`substring`}.
     * Optional Criteria: {`exclusive`}.
     */
    endsWith,

    /**
     * Label contains substring criterion.
     *
     * Required Criteria: {`substring`}.
     * Optional Criteria: {`exclusive`}.
     */
    contains,

    /**
     * Greater than comparator criterion.
     *
     * Required Criteria: {`comparator`}.
     */
    greaterThan,

    /**
     * Greater than or equal to comparator criterion.
     *
     * Required Criteria: {`comparator`}.
     */
    greaterThanOrEqualTo,

    /**
     * Less than comparator criterion.
     *
     * Required Criteria: {`comparator`}.
     */
    lessThan,

    /**
     * Less than or equal to comparator criterion.
     *
     * Required Criteria: {`comparator`}.
     */
    lessThanOrEqualTo,

    /**
     * Between `lowerBound` and `upperBound` criteria.
     *
     * Required Criteria: {`lowerBound`, `upperBound`}.
     * Optional Criteria: {`exclusive`}.
     */
    between,
}

/**
 * A simple enum that represents a type of filter for a PivotField.
 */
enum PivotFilterType {
    /**
     * `PivotFilterType` is unknown or unsupported.
     */
    unknown,

    /**
     * Filters based on the value of a PivotItem with respect to a `DataPivotHierarchy`.
     */
    value,

    /**
     * Filters specific manually selected PivotItems from the PivotTable.
     */
    manual,

    /**
     * Filters PivotItems based on their labels.
     * Note: A PivotField cannot simultaneously have a label filter and a date filter applied.
     */
    label,

    /**
     * Filters PivotItems with a date in place of a label.
     * Note: A PivotField cannot simultaneously have a label filter and a date filter applied.
     */
    date,
}

/**
 * A simple enum for top/bottom filters to select whether to filter by the top N or bottom N percent, number, or sum of values.
 */
enum TopBottomSelectionType {
    /**
     * Filter the top/bottom N number of items as measured by the chosen value.
     */
    items,

    /**
     * Filter the top/bottom N percent of items as measured by the chosen value.
     */
    percent,

    /**
     * Filter the top/bottom N sum as measured by the chosen value.
     */
    sum,
}

/**
 * Enum representing all accepted conditions by which a value filter can be applied.
 * Used to configure the type of PivotFilter that is applied to the field.
 * `PivotFilter.exclusive` can be set to `true` to invert many of these conditions.
 */
enum ValueFilterCondition {
    /**
     * `ValueFilterCondition` is unknown or unsupported.
     */
    unknown,

    /**
     * Equals comparator criterion.
     *
     * Required Criteria: {`value`, `comparator`}.
     * Optional Criteria: {`exclusive`}.
     */
    equals,

    /**
     * Greater than comparator criterion.
     *
     * Required Criteria: {`value`, `comparator`}.
     */
    greaterThan,

    /**
     * Greater than or equal to comparator criterion.
     *
     * Required Criteria: {`value`, `comparator`}.
     */
    greaterThanOrEqualTo,

    /**
     * Less than comparator criterion.
     *
     * Required Criteria: {`value`, `comparator`}.
     */
    lessThan,

    /**
     * Less than or equal to comparator criterion.
     *
     * Required Criteria: {`value`, `comparator`}.
     */
    lessThanOrEqualTo,

    /**
     * Between `lowerBound` and `upperBound` criteria.
     *
     * Required Criteria: {`value`, `lowerBound`, `upperBound`}.
     * Optional Criteria: {`exclusive`}.
     */
    between,

    /**
     * In top N (`threshold`) [items, percent, sum] of value category.
     *
     * Required Criteria: {`value`, `threshold`, `selectionType`}.
     */
    topN,

    /**
     * In bottom N (`threshold`) [items, percent, sum] of value category.
     *
     * Required Criteria: {`value`, `threshold`, `selectionType`}.
     */
    bottomN,
}

/**
 * Represents the dimensions when getting values from chart series.
 */
enum ChartSeriesDimension {
    /**
     * The chart series axis for the categories.
     */
    categories,

    /**
     * The chart series axis for the values.
     */
    values,

    /**
     * The chart series axis for the x-axis values in scatter and bubble charts.
     */
    xvalues,

    /**
     * The chart series axis for the y-axis values in scatter and bubble charts.
     */
    yvalues,

    /**
     * The chart series axis for the bubble sizes in bubble charts.
     */
    bubbleSizes,
}

/**
 * Represents the criteria for the top/bottom values filter.
 */
enum PivotFilterTopBottomCriterion {
    invalid,

    topItems,

    topPercent,

    topSum,

    bottomItems,

    bottomPercent,

    bottomSum,
}

/**
 * Represents the sort direction.
 */
enum SortBy {
    /**
     * Ascending sort. Smallest to largest or A to Z.
     */
    ascending,

    /**
     * Descending sort. Largest to smallest or Z to A.
     */
    descending,
}

/**
 * Aggregation function for the DataPivotField.
 */
enum AggregationFunction {
    /**
     * Aggregation function is unknown or unsupported.
     */
    unknown,

    /**
     * Excel will automatically select the aggregation based on the data items.
     */
    automatic,

    /**
     * Aggregate using the sum of the data, equivalent to the SUM function.
     */
    sum,

    /**
     * Aggregate using the count of items in the data, equivalent to the COUNTA function.
     */
    count,

    /**
     * Aggregate using the average of the data, equivalent to the AVERAGE function.
     */
    average,

    /**
     * Aggregate using the maximum value of the data, equivalent to the MAX function.
     */
    max,

    /**
     * Aggregate using the minimum value of the data, equivalent to the MIN function.
     */
    min,

    /**
     * Aggregate using the product of the data, equivalent to the PRODUCT function.
     */
    product,

    /**
     * Aggregate using the count of numbers in the data, equivalent to the COUNT function.
     */
    countNumbers,

    /**
     * Aggregate using the standard deviation of the data, equivalent to the STDEV function.
     */
    standardDeviation,

    /**
     * Aggregate using the standard deviation of the data, equivalent to the STDEVP function.
     */
    standardDeviationP,

    /**
     * Aggregate using the variance of the data, equivalent to the VAR function.
     */
    variance,

    /**
     * Aggregate using the variance of the data, equivalent to the VARP function.
     */
    varianceP,
}

/**
 * The ShowAs calculation function for the DataPivotField.
 */
enum ShowAsCalculation {
    /**
     * Calculation is unknown or unsupported.
     */
    unknown,

    /**
     * No calculation is applied.
     */
    none,

    /**
     * Percent of the grand total.
     */
    percentOfGrandTotal,

    /**
     * Percent of the row total.
     */
    percentOfRowTotal,

    /**
     * Percent of the column total.
     */
    percentOfColumnTotal,

    /**
     * Percent of the row total for the specified Base field.
     */
    percentOfParentRowTotal,

    /**
     * Percent of the column total for the specified Base field.
     */
    percentOfParentColumnTotal,

    /**
     * Percent of the grand total for the specified Base field.
     */
    percentOfParentTotal,

    /**
     * Percent of the specified Base field and Base item.
     */
    percentOf,

    /**
     * Running total of the specified Base field.
     */
    runningTotal,

    /**
     * Percent running total of the specified Base field.
     */
    percentRunningTotal,

    /**
     * Difference from the specified Base field and Base item.
     */
    differenceFrom,

    /**
     * Difference from the specified Base field and Base item.
     */
    percentDifferenceFrom,

    /**
     * Ascending rank of the specified Base field.
     */
    rankAscending,

    /**
     * Descending rank of the specified Base field.
     */
    rankDecending,

    /**
     * Calculates the values as follows:
     * ((value in cell) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total))
     */
    index,
}

/**
 * Represents the axis from which to get the PivotItems.
 */
enum PivotAxis {
    /**
     * The axis or region is unknown or unsupported.
     */
    unknown,

    /**
     * The row axis.
     */
    row,

    /**
     * The column axis.
     */
    column,

    /**
     * The data axis.
     */
    data,

    /**
     * The filter axis.
     */
    filter,
}

enum ChartAxisType {
    invalid,

    /**
     * Axis displays categories.
     */
    category,

    /**
     * Axis displays values.
     */
    value,

    /**
     * Axis displays data series.
     */
    series,
}

enum ChartAxisGroup {
    primary,

    secondary,
}

enum ChartAxisScaleType {
    linear,

    logarithmic,
}

enum ChartAxisPosition {
    automatic,

    maximum,

    minimum,

    custom,
}

enum ChartAxisTickMark {
    none,

    cross,

    inside,

    outside,
}

/**
 * Represents the state of calculation across the entire Excel application.
 */
enum CalculationState {
    /**
     * Calculations complete.
     */
    done,

    /**
     * Calculations in progress.
     */
    calculating,

    /**
     * Changes that trigger calculation have been made, but a recalculation has not yet been performed.
     */
    pending,
}

enum ChartAxisTickLabelPosition {
    nextToAxis,

    high,

    low,

    none,
}

enum ChartAxisDisplayUnit {
    /**
     * Default option. This will reset display unit to the axis, and set unit label invisible.
     */
    none,

    /**
     * This will set the axis in units of hundreds.
     */
    hundreds,

    /**
     * This will set the axis in units of thousands.
     */
    thousands,

    /**
     * This will set the axis in units of tens of thousands.
     */
    tenThousands,

    /**
     * This will set the axis in units of hundreds of thousands.
     */
    hundredThousands,

    /**
     * This will set the axis in units of millions.
     */
    millions,

    /**
     * This will set the axis in units of tens of millions.
     */
    tenMillions,

    /**
     * This will set the axis in units of hundreds of millions.
     */
    hundredMillions,

    /**
     * This will set the axis in units of billions.
     */
    billions,

    /**
     * This will set the axis in units of trillions.
     */
    trillions,

    /**
     * This will set the axis in units of custom value.
     */
    custom,
}

/**
 * Specifies the unit of time for chart axes and data series.
 */
enum ChartAxisTimeUnit {
    days,

    months,

    years,
}

/**
 * Represents the quartile calculation type of chart series layout. Only applies to a box and whisker chart.
 */
enum ChartBoxQuartileCalculation {
    inclusive,

    exclusive,
}

/**
 * Specifies the type of the category axis.
 */
enum ChartAxisCategoryType {
    /**
     * Excel controls the axis type.
     */
    automatic,

    /**
     * Axis groups data by an arbitrary set of categories.
     */
    textAxis,

    /**
     * Axis groups data on a time scale.
     */
    dateAxis,
}

/**
 * Specifies the bin type of a histogram chart or pareto chart series.
 */
enum ChartBinType {
    category,

    auto,

    binWidth,

    binCount,
}

enum ChartLineStyle {
    none,

    continuous,

    dash,

    dashDot,

    dashDotDot,

    dot,

    grey25,

    grey50,

    grey75,

    automatic,

    roundDot,
}

enum ChartDataLabelPosition {
    invalid,

    none,

    center,

    insideEnd,

    insideBase,

    outsideEnd,

    left,

    right,

    top,

    bottom,

    bestFit,

    callout,
}

/**
 * Represents which parts of the error bar to include.
 */
enum ChartErrorBarsInclude {
    both,

    minusValues,

    plusValues,
}

/**
 * Represents the range type for error bars.
 */
enum ChartErrorBarsType {
    fixedValue,

    percent,

    stDev,

    stError,

    custom,
}

/**
 * Represents the mapping level of a chart series. This only applies to region map charts.
 */
enum ChartMapAreaLevel {
    automatic,

    dataOnly,

    city,

    county,

    state,

    country,

    continent,

    world,
}

/**
 * Represents the gradient style of a chart series. This is only applicable for region map charts.
 */
enum ChartGradientStyle {
    twoPhaseColor,

    threePhaseColor,
}

/**
 * Represents the gradient style type of a chart series. This is only applicable for region map charts.
 */
enum ChartGradientStyleType {
    extremeValue,

    number,

    percent,
}

/**
 * Represents the position of the chart title.
 */
enum ChartTitlePosition {
    automatic,

    top,

    bottom,

    left,

    right,
}

enum ChartLegendPosition {
    invalid,

    top,

    bottom,

    left,

    right,

    corner,

    custom,
}

enum ChartMarkerStyle {
    invalid,

    automatic,

    none,

    square,

    diamond,

    triangle,

    x,

    star,

    dot,

    dash,

    circle,

    plus,

    picture,
}

enum ChartPlotAreaPosition {
    automatic,

    custom,
}

/**
 * Represents the region level of a chart series layout. This only applies to region map charts.
 */
enum ChartMapLabelStrategy {
    none,

    bestFit,

    showAll,
}

/**
 * Represents the region projection type of a chart series layout. This only applies to region map charts.
 */
enum ChartMapProjectionType {
    automatic,

    mercator,

    miller,

    robinson,

    albers,
}

/**
 * Represents the parent label strategy of the chart series layout. This only applies to treemap charts
 */
enum ChartParentLabelStrategy {
    none,

    banner,

    overlapping,
}

/**
 * Specifies whether the series are by rows or by columns. In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
 */
enum ChartSeriesBy {
    /**
     * In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
     */
    auto,

    columns,

    rows,
}

/**
 * Represents the horizontal alignment for the specified object.
 */
enum ChartTextHorizontalAlignment {
    center,

    left,

    right,

    justify,

    distributed,
}

/**
 * Represents the vertical alignment for the specified object.
 */
enum ChartTextVerticalAlignment {
    center,

    bottom,

    top,

    justify,

    distributed,
}

enum ChartTickLabelAlignment {
    center,

    left,

    right,
}

enum ChartType {
    invalid,

    columnClustered,

    columnStacked,

    columnStacked100,

    barClustered,

    barStacked,

    barStacked100,

    lineStacked,

    lineStacked100,

    lineMarkers,

    lineMarkersStacked,

    lineMarkersStacked100,

    pieOfPie,

    pieExploded,

    barOfPie,

    xyscatterSmooth,

    xyscatterSmoothNoMarkers,

    xyscatterLines,

    xyscatterLinesNoMarkers,

    areaStacked,

    areaStacked100,

    doughnutExploded,

    radarMarkers,

    radarFilled,

    surface,

    surfaceWireframe,

    surfaceTopView,

    surfaceTopViewWireframe,

    bubble,

    bubble3DEffect,

    stockHLC,

    stockOHLC,

    stockVHLC,

    stockVOHLC,

    cylinderColClustered,

    cylinderColStacked,

    cylinderColStacked100,

    cylinderBarClustered,

    cylinderBarStacked,

    cylinderBarStacked100,

    cylinderCol,

    coneColClustered,

    coneColStacked,

    coneColStacked100,

    coneBarClustered,

    coneBarStacked,

    coneBarStacked100,

    coneCol,

    pyramidColClustered,

    pyramidColStacked,

    pyramidColStacked100,

    pyramidBarClustered,

    pyramidBarStacked,

    pyramidBarStacked100,

    pyramidCol,

    line,

    pie,

    xyscatter,

    area,

    doughnut,

    radar,

    histogram,

    boxwhisker,

    pareto,

    regionMap,

    treemap,

    waterfall,

    sunburst,

    funnel,
}

enum ChartUnderlineStyle {
    none,

    single,
}

enum ChartDisplayBlanksAs {
    notPlotted,

    zero,

    interplotted,
}

enum ChartPlotBy {
    rows,

    columns,
}

enum ChartSplitType {
    splitByPosition,

    splitByValue,

    splitByPercentValue,

    splitByCustomSplit,
}

enum ChartColorScheme {
    colorfulPalette1,

    colorfulPalette2,

    colorfulPalette3,

    colorfulPalette4,

    monochromaticPalette1,

    monochromaticPalette2,

    monochromaticPalette3,

    monochromaticPalette4,

    monochromaticPalette5,

    monochromaticPalette6,

    monochromaticPalette7,

    monochromaticPalette8,

    monochromaticPalette9,

    monochromaticPalette10,

    monochromaticPalette11,

    monochromaticPalette12,

    monochromaticPalette13,
}

enum ChartTrendlineType {
    linear,

    exponential,

    logarithmic,

    movingAverage,

    polynomial,

    power,
}

/**
 * Specifies where in the z-order a shape should be moved relative to other shapes.
 */
enum ShapeZOrder {
    bringToFront,

    bringForward,

    sendToBack,

    sendBackward,
}

/**
 * Specifies the type of a shape.
 */
enum ShapeType {
    unsupported,

    image,

    geometricShape,

    group,

    line,
}

/**
 * Specifies whether the shape is scaled relative to its original or current size.
 */
enum ShapeScaleType {
    currentSize,

    originalSize,
}

/**
 * Specifies which part of the shape retains its position when the shape is scaled.
 */
enum ShapeScaleFrom {
    scaleFromTopLeft,

    scaleFromMiddle,

    scaleFromBottomRight,
}

/**
 * Specifies a shape's fill type.
 */
enum ShapeFillType {
    /**
     * No fill.
     */
    noFill,

    /**
     * Solid fill.
     */
    solid,

    /**
     * Gradient fill.
     */
    gradient,

    /**
     * Pattern fill.
     */
    pattern,

    /**
     * Picture and texture fill.
     */
    pictureAndTexture,

    /**
     * Mixed fill.
     */
    mixed,
}

/**
 * The type of underline applied to a font.
 */
enum ShapeFontUnderlineStyle {
    none,

    single,

    double,

    heavy,

    dotted,

    dottedHeavy,

    dash,

    dashHeavy,

    dashLong,

    dashLongHeavy,

    dotDash,

    dotDashHeavy,

    dotDotDash,

    dotDotDashHeavy,

    wavy,

    wavyHeavy,

    wavyDouble,
}

/**
 * The format of the image.
 */
enum PictureFormat {
    unknown,

    /**
     * Bitmap image.
     */
    bmp,

    /**
     * Joint Photographic Experts Group.
     */
    jpeg,

    /**
     * Graphics Interchange Format.
     */
    gif,

    /**
     * Portable Network Graphics.
     */
    png,

    /**
     * Scalable Vector Graphic.
     */
    svg,
}

/**
 * The style for a line.
 */
enum ShapeLineStyle {
    /**
     * Single line.
     */
    single,

    /**
     * Thick line with a thin line on each side.
     */
    thickBetweenThin,

    /**
     * Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line.
     */
    thickThin,

    /**
     * Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line.
     */
    thinThick,

    /**
     * Two thin lines.
     */
    thinThin,
}

/**
 * The dash style for a line.
 */
enum ShapeLineDashStyle {
    dash,

    dashDot,

    dashDotDot,

    longDash,

    longDashDot,

    roundDot,

    solid,

    squareDot,

    longDashDotDot,

    systemDash,

    systemDot,

    systemDashDot,
}

enum ArrowheadLength {
    short,

    medium,

    long,
}

enum ArrowheadStyle {
    none,

    triangle,

    stealth,

    diamond,

    oval,

    open,
}

enum ArrowheadWidth {
    narrow,

    medium,

    wide,
}

enum BindingType {
    range,

    table,

    text,
}

enum BorderIndex {
    edgeTop,

    edgeBottom,

    edgeLeft,

    edgeRight,

    insideVertical,

    insideHorizontal,

    diagonalDown,

    diagonalUp,
}

enum BorderLineStyle {
    none,

    continuous,

    dash,

    dashDot,

    dashDotDot,

    dot,

    double,

    slantDashDot,
}

enum BorderWeight {
    hairline,

    thin,

    medium,

    thick,
}

enum CalculationMode {
    /**
     * The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.
     */
    automatic,

    /**
     * Calculates new formula results every time the relevant data is changed, unless the formula is in a data table.
     */
    automaticExceptTables,

    /**
     * Calculations only occur when the user or add-in requests them.
     */
    manual,
}

enum CalculationType {
    /**
     * Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty.
     */
    recalculate,

    /**
     * This will mark all cells as dirty and then recalculate them.
     */
    full,

    /**
     * This will rebuild the full dependency chain, mark all cells as dirty and then recalculate them.
     */
    fullRebuild,
}

enum ClearApplyTo {
    all,

    /**
     * Clears all formatting for the range.
     */
    formats,

    /**
     * Clears the contents of the range.
     */
    contents,

    /**
     * Clears all hyperlinks, but leaves all content and formatting intact.
     */
    hyperlinks,

    /**
     * Removes hyperlinks and formatting for the cell but leaves content, conditional formats, and data validation intact.
     */
    removeHyperlinks,
}

/**
 * Represents the format options for a data bar axis.
 */
enum ConditionalDataBarAxisFormat {
    automatic,

    none,

    cellMidPoint,
}

/**
 * Represents the data bar direction within a cell.
 */
enum ConditionalDataBarDirection {
    context,

    leftToRight,

    rightToLeft,
}

/**
 * Represents the direction for a selection.
 */
enum ConditionalFormatDirection {
    top,

    bottom,
}

enum ConditionalFormatType {
    custom,

    dataBar,

    colorScale,

    iconSet,

    topBottom,

    presetCriteria,

    containsText,

    cellValue,
}

/**
 * Represents the types of conditional format values.
 */
enum ConditionalFormatRuleType {
    invalid,

    automatic,

    lowestValue,

    highestValue,

    number,

    percent,

    formula,

    percentile,
}

/**
 * Represents the types of icon conditional format.
 */
enum ConditionalFormatIconRuleType {
    invalid,

    number,

    percent,

    formula,

    percentile,
}

/**
 * Represents the types of color criterion for conditional formatting.
 */
enum ConditionalFormatColorCriterionType {
    invalid,

    lowestValue,

    highestValue,

    number,

    percent,

    formula,

    percentile,
}

/**
 * Represents the criteria for the above/below average conditional format type.
 */
enum ConditionalTopBottomCriterionType {
    invalid,

    topItems,

    topPercent,

    bottomItems,

    bottomPercent,
}

/**
 * Represents the criteria of the preset criteria conditional format type.
 */
enum ConditionalFormatPresetCriterion {
    invalid,

    blanks,

    nonBlanks,

    errors,

    nonErrors,

    yesterday,

    today,

    tomorrow,

    lastSevenDays,

    lastWeek,

    thisWeek,

    nextWeek,

    lastMonth,

    thisMonth,

    nextMonth,

    aboveAverage,

    belowAverage,

    equalOrAboveAverage,

    equalOrBelowAverage,

    oneStdDevAboveAverage,

    oneStdDevBelowAverage,

    twoStdDevAboveAverage,

    twoStdDevBelowAverage,

    threeStdDevAboveAverage,

    threeStdDevBelowAverage,

    uniqueValues,

    duplicateValues,
}

/**
 * Represents the operator of the text conditional format type.
 */
enum ConditionalTextOperator {
    invalid,

    contains,

    notContains,

    beginsWith,

    endsWith,
}

/**
 * Represents the operator of the text conditional format type.
 */
enum ConditionalCellValueOperator {
    invalid,

    between,

    notBetween,

    equalTo,

    notEqualTo,

    greaterThan,

    lessThan,

    greaterThanOrEqual,

    lessThanOrEqual,
}

/**
 * Represents the operator for each icon criteria.
 */
enum ConditionalIconCriterionOperator {
    invalid,

    greaterThan,

    greaterThanOrEqual,
}

enum ConditionalRangeBorderIndex {
    edgeTop,

    edgeBottom,

    edgeLeft,

    edgeRight,
}

enum ConditionalRangeBorderLineStyle {
    none,

    continuous,

    dash,

    dashDot,

    dashDotDot,

    dot,
}

enum ConditionalRangeFontUnderlineStyle {
    none,

    single,

    double,
}

/**
 * Represents the data validation type enum.
 */
enum DataValidationType {
    /**
     * None means allow any value, indicating that there is no data validation in the range.
     */
    none,

    /**
     * The whole number data validation type.
     */
    wholeNumber,

    /**
     * The decimal data validation type.
     */
    decimal,

    /**
     * The list data validation type.
     */
    list,

    /**
     * The date data validation type.
     */
    date,

    /**
     * The time data validation type.
     */
    time,

    /**
     * The text length data validation type.
     */
    textLength,

    /**
     * The custom data validation type.
     */
    custom,

    /**
     * Inconsistent means that the range has inconsistent data validation, indicating that there are different rules on different cells.
     */
    inconsistent,

    /**
     * Mixed criteria means that the range has data validation present on some but not all cells.
     */
    mixedCriteria,
}

/**
 * Represents the data validation operator enum.
 */
enum DataValidationOperator {
    between,

    notBetween,

    equalTo,

    notEqualTo,

    greaterThan,

    lessThan,

    greaterThanOrEqualTo,

    lessThanOrEqualTo,
}

/**
 * Represents the data validation error alert style. The default is `Stop`.
 */
enum DataValidationAlertStyle {
    stop,

    warning,

    information,
}

enum DeleteShiftDirection {
    up,

    left,
}

enum DynamicFilterCriteria {
    unknown,

    aboveAverage,

    allDatesInPeriodApril,

    allDatesInPeriodAugust,

    allDatesInPeriodDecember,

    allDatesInPeriodFebruary,

    allDatesInPeriodJanuary,

    allDatesInPeriodJuly,

    allDatesInPeriodJune,

    allDatesInPeriodMarch,

    allDatesInPeriodMay,

    allDatesInPeriodNovember,

    allDatesInPeriodOctober,

    allDatesInPeriodQuarter1,

    allDatesInPeriodQuarter2,

    allDatesInPeriodQuarter3,

    allDatesInPeriodQuarter4,

    allDatesInPeriodSeptember,

    belowAverage,

    lastMonth,

    lastQuarter,

    lastWeek,

    lastYear,

    nextMonth,

    nextQuarter,

    nextWeek,

    nextYear,

    thisMonth,

    thisQuarter,

    thisWeek,

    thisYear,

    today,

    tomorrow,

    yearToDate,

    yesterday,
}

enum FilterDatetimeSpecificity {
    year,

    month,

    day,

    hour,

    minute,

    second,
}

enum FilterOn {
    bottomItems,

    bottomPercent,

    cellColor,

    dynamic,

    fontColor,

    values,

    topItems,

    topPercent,

    icon,

    custom,
}

enum FilterOperator {
    and,

    or,
}

enum HorizontalAlignment {
    general,

    left,

    center,

    right,

    fill,

    justify,

    centerAcrossSelection,

    distributed,
}

enum IconSet {
    invalid,

    threeArrows,

    threeArrowsGray,

    threeFlags,

    threeTrafficLights1,

    threeTrafficLights2,

    threeSigns,

    threeSymbols,

    threeSymbols2,

    fourArrows,

    fourArrowsGray,

    fourRedToBlack,

    fourRating,

    fourTrafficLights,

    fiveArrows,

    fiveArrowsGray,

    fiveRating,

    fiveQuarters,

    threeStars,

    threeTriangles,

    fiveBoxes,
}

enum ImageFittingMode {
    fit,

    fitAndCenter,

    fill,
}

enum InsertShiftDirection {
    down,

    right,
}

enum NamedItemScope {
    worksheet,

    workbook,
}

enum NamedItemType {
    string,

    integer,

    double,

    boolean,

    range,

    error,

    array,
}

enum RangeUnderlineStyle {
    none,

    single,

    double,

    singleAccountant,

    doubleAccountant,
}

enum SheetVisibility {
    visible,

    hidden,

    veryHidden,
}

enum RangeValueType {
    unknown,

    empty,

    string,

    integer,

    double,

    boolean,

    error,

    richValue,
}

/**
 * Specifies the search direction.
 */
enum SearchDirection {
    /**
     * Search in forward order.
     */
    forward,

    /**
     * Search in reverse order.
     */
    backwards,
}

enum SortOrientation {
    rows,

    columns,
}

enum SortOn {
    value,

    cellColor,

    fontColor,

    icon,
}

enum SortDataOption {
    normal,

    textAsNumber,
}

enum SortMethod {
    pinYin,

    strokeCount,
}

enum VerticalAlignment {
    top,

    center,

    bottom,

    justify,

    distributed,
}

enum DocumentPropertyType {
    number,

    boolean,

    date,

    string,

    float,
}

enum SubtotalLocationType {
    /**
     * Subtotals are at the top.
     */
    atTop,

    /**
     * Subtotals are at the bottom.
     */
    atBottom,

    /**
     * Subtotals are off.
     */
    off,
}

enum PivotLayoutType {
    /**
     * A horizontally compressed form with labels from the next field in the same column.
     */
    compact,

    /**
     * Inner fields' items are always on a new line relative to the outer fields' items.
     */
    tabular,

    /**
     * Inner fields' items are on same row as outer fields' items and subtotals are always on the bottom.
     */
    outline,
}

enum ProtectionSelectionMode {
    /**
     * Selection is allowed for all cells.
     */
    normal,

    /**
     * Selection is allowed only for cells that are not locked.
     */
    unlocked,

    /**
     * Selection is not allowed for any cells.
     */
    none,
}

enum PageOrientation {
    portrait,

    landscape,
}

enum PaperType {
    letter,

    letterSmall,

    tabloid,

    ledger,

    legal,

    statement,

    executive,

    a3,

    a4,

    a4Small,

    a5,

    b4,

    b5,

    folio,

    quatro,

    paper10x14,

    paper11x17,

    note,

    envelope9,

    envelope10,

    envelope11,

    envelope12,

    envelope14,

    csheet,

    dsheet,

    esheet,

    envelopeDL,

    envelopeC5,

    envelopeC3,

    envelopeC4,

    envelopeC6,

    envelopeC65,

    envelopeB4,

    envelopeB5,

    envelopeB6,

    envelopeItaly,

    envelopeMonarch,

    envelopePersonal,

    fanfoldUS,

    fanfoldStdGerman,

    fanfoldLegalGerman,
}

enum ReadingOrder {
    /**
     * Reading order is determined by the language of the first character entered.
     * If a right-to-left language character is entered first, reading order is right to left.
     * If a left-to-right language character is entered first, reading order is left to right.
     */
    context,

    /**
     * Left to right reading order
     */
    leftToRight,

    /**
     * Right to left reading order
     */
    rightToLeft,
}

enum BuiltInStyle {
    normal,

    comma,

    currency,

    percent,

    wholeComma,

    wholeDollar,

    hlink,

    hlinkTrav,

    note,

    warningText,

    emphasis1,

    emphasis2,

    emphasis3,

    sheetTitle,

    heading1,

    heading2,

    heading3,

    heading4,

    input,

    output,

    calculation,

    checkCell,

    linkedCell,

    total,

    good,

    bad,

    neutral,

    accent1,

    accent1_20,

    accent1_40,

    accent1_60,

    accent2,

    accent2_20,

    accent2_40,

    accent2_60,

    accent3,

    accent3_20,

    accent3_40,

    accent3_60,

    accent4,

    accent4_20,

    accent4_40,

    accent4_60,

    accent5,

    accent5_20,

    accent5_40,

    accent5_60,

    accent6,

    accent6_20,

    accent6_40,

    accent6_60,

    explanatoryText,
}

enum PrintErrorType {
    asDisplayed,

    blank,

    dash,

    notAvailable,
}

enum WorksheetPositionType {
    none,

    before,

    after,

    beginning,

    end,
}

enum PrintComments {
    /**
     * Comments will not be printed.
     */
    noComments,

    /**
     * Comments will be printed as end notes at the end of the worksheet.
     */
    endSheet,

    /**
     * Comments will be printed where they were inserted in the worksheet.
     */
    inPlace,
}

enum PrintOrder {
    /**
     * Process down the rows before processing across pages or page fields to the right.
     */
    downThenOver,

    /**
     * Process across pages or page fields to the right before moving down the rows.
     */
    overThenDown,
}

enum PrintMarginUnit {
    /**
     * Assign the page margins in points. A point is 1/72 of an inch.
     */
    points,

    /**
     * Assign the page margins in inches.
     */
    inches,

    /**
     * Assign the page margins in centimeters.
     */
    centimeters,
}

enum HeaderFooterState {
    /**
     * Only one general header/footer is used for all pages printed.
     */
    default,

    /**
     * There is a separate first page header/footer, and a general header/footer used for all other pages.
     */
    firstAndDefault,

    /**
     * There is a different header/footer for odd and even pages.
     */
    oddAndEven,

    /**
     * There is a separate first page header/footer, then there is a separate header/footer for odd and even pages.
     */
    firstOddAndEven,
}

/**
 * The behavior types when AutoFill is used on a range in the workbook.
 */
enum AutoFillType {
    /**
     * Populates the adjacent cells based on the surrounding data (the standard AutoFill behavior).
     */
    fillDefault,

    /**
     * Populates the adjacent cells with data based on the selected data.
     */
    fillCopy,

    /**
     * Populates the adjacent cells with data that follows a pattern in the copied cells.
     */
    fillSeries,

    /**
     * Populates the adjacent cells with the selected formulas.
     */
    fillFormats,

    /**
     * Populates the adjacent cells with the selected values.
     */
    fillValues,

    /**
     * A version of "FillSeries" for dates that bases the pattern on either the day of the month or the day of the week, depending on the context.
     */
    fillDays,

    /**
     * A version of "FillSeries" for dates that bases the pattern on the day of the week and only includes weekdays.
     */
    fillWeekdays,

    /**
     * A version of "FillSeries" for dates that bases the pattern on the month.
     */
    fillMonths,

    /**
     * A version of "FillSeries" for dates that bases the pattern on the year.
     */
    fillYears,

    /**
     * A version of "FillSeries" for numbers that fills out the values in the adjacent cells according to a linear trend model.
     */
    linearTrend,

    /**
     * A version of "FillSeries" for numbers that fills out the values in the adjacent cells according to a growth trend model.
     */
    growthTrend,

    /**
     * Populates the adjacent cells by using Excel's Flash Fill feature.
     */
    flashFill,
}

enum GroupOption {
    /**
     * Group by rows.
     */
    byRows,

    /**
     * Group by columns.
     */
    byColumns,
}

enum RangeCopyType {
    all,

    formulas,

    values,

    formats,
}

enum LinkedDataTypeState {
    none,

    validLinkedData,

    disambiguationNeeded,

    brokenLinkedData,

    fetchingData,
}

/**
 * Specifies the shape type for a `GeometricShape` object.
 */
enum GeometricShapeType {
    lineInverse,

    triangle,

    rightTriangle,

    rectangle,

    diamond,

    parallelogram,

    trapezoid,

    nonIsoscelesTrapezoid,

    pentagon,

    hexagon,

    heptagon,

    octagon,

    decagon,

    dodecagon,

    star4,

    star5,

    star6,

    star7,

    star8,

    star10,

    star12,

    star16,

    star24,

    star32,

    roundRectangle,

    round1Rectangle,

    round2SameRectangle,

    round2DiagonalRectangle,

    snipRoundRectangle,

    snip1Rectangle,

    snip2SameRectangle,

    snip2DiagonalRectangle,

    plaque,

    ellipse,

    teardrop,

    homePlate,

    chevron,

    pieWedge,

    pie,

    blockArc,

    donut,

    noSmoking,

    rightArrow,

    leftArrow,

    upArrow,

    downArrow,

    stripedRightArrow,

    notchedRightArrow,

    bentUpArrow,

    leftRightArrow,

    upDownArrow,

    leftUpArrow,

    leftRightUpArrow,

    quadArrow,

    leftArrowCallout,

    rightArrowCallout,

    upArrowCallout,

    downArrowCallout,

    leftRightArrowCallout,

    upDownArrowCallout,

    quadArrowCallout,

    bentArrow,

    uturnArrow,

    circularArrow,

    leftCircularArrow,

    leftRightCircularArrow,

    curvedRightArrow,

    curvedLeftArrow,

    curvedUpArrow,

    curvedDownArrow,

    swooshArrow,

    cube,

    can,

    lightningBolt,

    heart,

    sun,

    moon,

    smileyFace,

    irregularSeal1,

    irregularSeal2,

    foldedCorner,

    bevel,

    frame,

    halfFrame,

    corner,

    diagonalStripe,

    chord,

    arc,

    leftBracket,

    rightBracket,

    leftBrace,

    rightBrace,

    bracketPair,

    bracePair,

    callout1,

    callout2,

    callout3,

    accentCallout1,

    accentCallout2,

    accentCallout3,

    borderCallout1,

    borderCallout2,

    borderCallout3,

    accentBorderCallout1,

    accentBorderCallout2,

    accentBorderCallout3,

    wedgeRectCallout,

    wedgeRRectCallout,

    wedgeEllipseCallout,

    cloudCallout,

    cloud,

    ribbon,

    ribbon2,

    ellipseRibbon,

    ellipseRibbon2,

    leftRightRibbon,

    verticalScroll,

    horizontalScroll,

    wave,

    doubleWave,

    plus,

    flowChartProcess,

    flowChartDecision,

    flowChartInputOutput,

    flowChartPredefinedProcess,

    flowChartInternalStorage,

    flowChartDocument,

    flowChartMultidocument,

    flowChartTerminator,

    flowChartPreparation,

    flowChartManualInput,

    flowChartManualOperation,

    flowChartConnector,

    flowChartPunchedCard,

    flowChartPunchedTape,

    flowChartSummingJunction,

    flowChartOr,

    flowChartCollate,

    flowChartSort,

    flowChartExtract,

    flowChartMerge,

    flowChartOfflineStorage,

    flowChartOnlineStorage,

    flowChartMagneticTape,

    flowChartMagneticDisk,

    flowChartMagneticDrum,

    flowChartDisplay,

    flowChartDelay,

    flowChartAlternateProcess,

    flowChartOffpageConnector,

    actionButtonBlank,

    actionButtonHome,

    actionButtonHelp,

    actionButtonInformation,

    actionButtonForwardNext,

    actionButtonBackPrevious,

    actionButtonEnd,

    actionButtonBeginning,

    actionButtonReturn,

    actionButtonDocument,

    actionButtonSound,

    actionButtonMovie,

    gear6,

    gear9,

    funnel,

    mathPlus,

    mathMinus,

    mathMultiply,

    mathDivide,

    mathEqual,

    mathNotEqual,

    cornerTabs,

    squareTabs,

    plaqueTabs,

    chartX,

    chartStar,

    chartPlus,
}

enum ConnectorType {
    straight,

    elbow,

    curve,
}

enum ContentType {
    /**
     * Indicates a plain format type for the comment content.
     */
    plain,

    /**
     * Comment content containing mentions.
     */
    mention,
}

enum SpecialCellType {
    /**
     * All cells with conditional formats.
     */
    conditionalFormats,

    /**
     * Cells with validation criteria.
     */
    dataValidations,

    /**
     * Cells with no content.
     */
    blanks,

    /**
     * Cells containing constants.
     */
    constants,

    /**
     * Cells containing formulas.
     */
    formulas,

    /**
     * Cells with the same conditional format as the first cell in the range.
     */
    sameConditionalFormat,

    /**
     * Cells with the same data validation criteria as the first cell in the range.
     */
    sameDataValidation,

    /**
     * Cells that are visible.
     */
    visible,
}

enum SpecialCellValueType {
    /**
     * Cells that have errors, boolean, numeric, or string values.
     */
    all,

    /**
     * Cells that have errors.
     */
    errors,

    /**
     * Cells that have errors or boolean values.
     */
    errorsLogical,

    /**
     * Cells that have errors or numeric values.
     */
    errorsNumbers,

    /**
     * Cells that have errors or string values.
     */
    errorsText,

    /**
     * Cells that have errors, boolean, or numeric values.
     */
    errorsLogicalNumber,

    /**
     * Cells that have errors, boolean, or string values.
     */
    errorsLogicalText,

    /**
     * Cells that have errors, numeric, or string values.
     */
    errorsNumberText,

    /**
     * Cells that have a boolean value.
     */
    logical,

    /**
     * Cells that have a boolean or numeric value.
     */
    logicalNumbers,

    /**
     * Cells that have a boolean or string value.
     */
    logicalText,

    /**
     * Cells that have a boolean, numeric, or string value.
     */
    logicalNumbersText,

    /**
     * Cells that have a numeric value.
     */
    numbers,

    /**
     * Cells that have a numeric or string value.
     */
    numbersText,

    /**
     * Cells that have a string value.
     */
    text,
}

/**
 * Specifies the way that an object is attached to its underlying cells.
 */
enum Placement {
    /**
     * The object is moved with the cells.
     */
    twoCell,

    /**
     * The object is moved and sized with the cells.
     */
    oneCell,

    /**
     * The object is free floating.
     */
    absolute,
}

enum FillPattern {
    none,

    solid,

    gray50,

    gray75,

    gray25,

    horizontal,

    vertical,

    down,

    up,

    checker,

    semiGray75,

    lightHorizontal,

    lightVertical,

    lightDown,

    lightUp,

    grid,

    crissCross,

    gray16,

    gray8,

    linearGradient,

    rectangularGradient,
}

/**
 * Specifies the horizontal alignment for the text frame in a shape.
 */
enum ShapeTextHorizontalAlignment {
    left,

    center,

    right,

    justify,

    justifyLow,

    distributed,

    thaiDistributed,
}

/**
 * Specifies the vertical alignment for the text frame in a shape.
 */
enum ShapeTextVerticalAlignment {
    top,

    middle,

    bottom,

    justified,

    distributed,
}

/**
 * Specifies the vertical overflow for the text frame in a shape.
 */
enum ShapeTextVerticalOverflow {
    /**
     * Allow text to overflow the text frame vertically (can be from the top, bottom, or both depending on the text alignment).
     */
    overflow,

    /**
     * Hide text that does not fit vertically within the text frame, and add an ellipsis (...) at the end of the visible text.
     */
    ellipsis,

    /**
     * Hide text that does not fit vertically within the text frame.
     */
    clip,
}

/**
 * Specifies the horizontal overflow for the text frame in a shape.
 */
enum ShapeTextHorizontalOverflow {
    overflow,

    clip,
}

/**
 * Specifies the reading order for the text frame in a shape.
 */
enum ShapeTextReadingOrder {
    leftToRight,

    rightToLeft,
}

/**
 * Specifies the orientation for the text frame in a shape.
 */
enum ShapeTextOrientation {
    horizontal,

    vertical,

    vertical270,

    wordArtVertical,

    eastAsianVertical,

    mongolianVertical,

    wordArtVerticalRTL,
}

/**
 * Determines the type of automatic sizing allowed.
 */
enum ShapeAutoSize {
    /**
     * No autosizing.
     */
    autoSizeNone,

    /**
     * The text is adjusted to fit the shape.
     */
    autoSizeTextToFitShape,

    /**
     * The shape is adjusted to fit the text.
     */
    autoSizeShapeToFitText,

    /**
     * A combination of automatic sizing schemes are used.
     */
    autoSizeMixed,
}

/**
 * Specifies the slicer sort behavior for `Slicer.sortBy`.
 */
enum SlicerSortType {
    /**
     * Sort slicer items in the order provided by the data source.
     */
    dataSourceOrder,

    /**
     * Sort slicer items in ascending order by item captions.
     */
    ascending,

    /**
     * Sort slicer items in descending order by item captions.
     */
    descending,
}

/**
 * Represents a category of number formats.
 */
enum NumberFormatCategory {
    /**
     * General format cells have no specific number format.
     */
    general,

    /**
     * Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.
     */
    number,

    /**
     * Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.
     */
    currency,

    /**
     * Accounting formats line up the currency symbols and decimal points in a column.
     */
    accounting,

    /**
     * Date formats display date and time serial numbers as date values. Date formats that begin with an asterisk (*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings.
     */
    date,

    /**
     * Time formats display date and time serial numbers as date values. Time formats that begin with an asterisk (*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings.
     */
    time,

    /**
     * Percentage formats multiply the cell value by 100 and displays the result with a percent symbol.
     */
    percentage,

    /**
     * Fraction formats display the cell value as a whole number with the remainder rounded to the nearest fraction value.
     */
    fraction,

    /**
     * Scientific formats display the cell value as a number between 1 and 10 multiplied by a power of 10.
     */
    scientific,

    /**
     * Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.
     */
    text,

    /**
     * Special formats are useful for tracking list and database values.
     */
    special,

    /**
     * A custom format that is not a part of any category.
     */
    custom,
}


