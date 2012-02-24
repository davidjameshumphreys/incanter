(ns
    ^{:doc "Functions for reading and writing to cells."}
  incanter.excel.cells
  (:import [org.apache.poi.ss.usermodel Cell CellStyle DateUtil CellValue]
           [org.apache.poi.ss.usermodel Row Sheet]))
(set! *warn-on-reflection* true)

;; Changed from a multi-method to remove reflection warnings.
(defn ^Cell write-cell-value [^Cell cell ^Object i]
  (do
    (cond
     (isa? (. i getClass) Number) (.setCellValue cell (double (. ^Number i doubleValue)))
     (keyword? i)                 (.setCellValue cell (str (name i)))
     :else                        (.setCellValue cell (str i)))
    cell))

(defmulti
  get-cell-formula-value
  "Get the value after the evaluating the formula.  See http://poi.apache.org/spreadsheet/eval.html#Evaluate"
  (fn [^Cell evaled-cell evaled-type]
    evaled-type))

(defmulti
  get-cell-value
  "Get the cell value depending on the cell type."
  (fn [^CellValue cell]
    (let [ct (. cell getCellType)]
      (if (not (= Cell/CELL_TYPE_NUMERIC ct))
        ct
        (if (DateUtil/isCellDateFormatted cell)
          :date
          ct)))))

(defn write-line [^Sheet sheet row-num line ^CellStyle style]
  (let [^Row xl-line (. sheet createRow row-num)]
    (dorun
     (map
      #(if (not (nil? %2))
         (.
          (write-cell-value (. xl-line createCell %1) %2)
          setCellStyle style))
      (iterate inc 0)
      (seq line)))))

(defn write-line-values [{:keys [sheet bold normal workbook]} bold? row-number coll]
  (write-line sheet row-number coll (if bold? bold normal)))
(defn cell-iterator [^Row row]
  (if row
    (for [idx (range (.getFirstCellNum row) (.getLastCellNum row))]
      (if-let [cell (.getCell row idx)]
        cell
        (.createCell row idx Cell/CELL_TYPE_BLANK)))
    ()))

(defn read-line-values [row-iterator-item]
  (doall (map get-cell-value (cell-iterator row-iterator-item))))

;; Implementations of the multi-methods:

(defmethod get-cell-formula-value
  Cell/CELL_TYPE_BOOLEAN [^CellValue evaled-cell evaled-type]
  (. evaled-cell getBooleanValue))

(defmethod get-cell-formula-value
  Cell/CELL_TYPE_STRING  [^CellValue evaled-cell evaled-type]
  (. evaled-cell getStringValue))

(defmethod get-cell-formula-value
  :number                [^CellValue evaled-cell evaled-type]
  (. evaled-cell getNumberValue))

(defmethod get-cell-formula-value
  :date                  [^CellValue evaled-cell evaled-type]
  (DateUtil/getJavaDate (. evaled-cell getNumberValue)))

(defmethod get-cell-formula-value
  :default               [^CellValue evaled-cell evaled-type]
  (str "Unknown cell type " (. evaled-cell getCellType)))

(defmethod get-cell-value Cell/CELL_TYPE_BLANK   [^CellValue cell])
(defmethod get-cell-value Cell/CELL_TYPE_FORMULA [^CellValue cell]
  (let [^CellValue val (.
             (.. cell
                 getSheet
                 getWorkbook
                 getCreationHelper
                 createFormulaEvaluator)
             evaluate cell)
        evaluated-type (. val getCellType)]
    (get-cell-formula-value
     val
     (if (= Cell/CELL_TYPE_NUMERIC evaluated-type)
       (if (DateUtil/isCellInternalDateFormatted cell)
         ;; Check the original for date formatting hints
         :date
         :number)
       evaluated-type))))

(defmethod get-cell-value Cell/CELL_TYPE_BOOLEAN [^Cell cell]
  (. cell getBooleanCellValue))
(defmethod get-cell-value Cell/CELL_TYPE_STRING  [^Cell cell]
  (. cell getStringCellValue))
(defmethod get-cell-value Cell/CELL_TYPE_NUMERIC [^Cell cell]
  (. cell getNumericCellValue))
(defmethod get-cell-value :date [^Cell cell]
  (. cell getDateCellValue))
(defmethod get-cell-value :default [^Cell cell]
  (str "Unknown cell type " (. cell getCellType)))

