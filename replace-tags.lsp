; replace-tags.lsp
; AutoLISP script to replace text to the right of tags based on Excel or CSV file
; Usage: load this LISP in AutoCAD (appload or drag), then run REPLTAG command
; Input: Excel (.xlsx/.xls) or CSV file where column 1 = tag text, column 2 = replacement text
; Behavior: for each tag occurrence, finds the nearest text/mtext/attribute to the right
; and replaces its content with the replacement value.
; 
; Notes:
; - Supports both Excel files (via COM) and CSV files (for portability).
; - Excel COM requires full AutoCAD and appropriate system permissions.
; - Processes ALL tag/replacement pairs from the file automatically.
; - Provides detailed progress output and comprehensive summary statistics.
;
; Latest version: Rewritten with modern AutoLISP practices including:
; - Enhanced error handling with actionable error messages
; - Modular function design for better maintainability
; - Comprehensive input validation and user feedback
; - Fixed critical bug where file contents were ignored

(vl-load-com)

(defun split (s sep / res pos)
  (if (or (not s) (equal s ""))
    '()
    (progn
      (setq res '())
      (while (setq pos (vl-string-search sep s))
        (setq res (cons (substr s 1 pos) res))
        (setq s (substr s (+ pos (strlen sep)) ))
      )
      (setq res (cons s res))
      (reverse res)
    )
  )
)

(defun trim (s / len start end)
  (if (not s) ""
    (progn
      (setq len (strlen s) start 1 end len)
      (while (and (<= start len) (member (substr s start 1) '(" " "\t" "\n" "\r")))
        (setq start (1+ start))
      )
      (while (and (>= end start) (member (substr s end 1) '(" " "\t" "\n" "\r")))
        (setq end (1- end))
      )
      (if (> end (1- start)) (substr s start (- end start -1)) ""
    )
  )
)

(defun read-csv-pairs (file / fh line parts tag repl pairs line-num)
  (setq pairs '() line-num 0)
  (princ (strcat "\nReading CSV file: " file))

  (if (not (setq fh (open file "r")))
    (princ (strcat "\n*** ERROR: Could not open CSV file: " file))
    (progn
      (princ "\n✓ CSV file opened successfully")
      (while (setq line (read-line fh))
        (setq line-num (1+ line-num))
        (if (not (equal (trim line) ""))
          (progn
            (setq parts (split line ","))
            (if (< (length parts) 2)
              (princ (strcat "\n  Warning: Line " (itoa line-num) " has less than 2 columns, skipping"))
              (progn
                (setq tag (trim (nth 0 parts)))
                (setq repl (trim (nth 1 parts)))
                (if (and (not (equal tag "")) (not (equal repl "")))
                  (setq pairs (cons (list tag repl) pairs))
                  (princ (strcat "\n  Warning: Line " (itoa line-num) " has empty tag or replacement, skipping"))
                )
              )
            )
          )
        )
      )
      (close fh)
      (princ (strcat "\n✓ Successfully loaded " (itoa (length pairs)) " valid pairs from CSV"))
    )
  )
  (reverse pairs)
)

(defun read-excel-pairs (file / xl wb ws usedrange rows cols pairs row-count col-count row cell-tag cell-repl tag repl error-occurred err-obj)
  (setq pairs '() error-occurred nil)

  ; Try to create Excel COM object with better error handling
  (setq err-obj (vl-catch-all-apply 'vlax-get-or-create-object '("Excel.Application")))
  (if (vl-catch-all-error-p err-obj)
    (progn
      (princ "\n*** ERROR: Excel COM not available.")
      (princ "\nPossible reasons:")
      (princ "\n  - Microsoft Excel is not installed")
      (princ "\n  - AutoCAD doesn't have permission to launch COM applications")
      (princ "\n  - Excel COM is blocked by security policies")
      (princ "\nSolution: Please use CSV format instead.")
      (setq error-occurred T)
    )
    (progn
      (setq xl err-obj)
      (princ "\n✓ Excel COM object created successfully")
      (princ (strcat "\nOpening workbook: " file))

      ; Try to open the workbook
      (setq err-obj (vl-catch-all-apply 'vlax-invoke-method
                     (list (vlax-get-property xl 'Workbooks) "Open" file)))
      (if (vl-catch-all-error-p err-obj)
        (progn
          (princ (strcat "\n*** ERROR: Could not open Excel file: " file))
          (princ "\nPossible reasons:")
          (princ "\n  - File is corrupted or in use by another program")
          (princ "\n  - File path is incorrect")
          (princ "\n  - File format is not supported")
          (if xl (vlax-release-object xl))
          (setq error-occurred T)
        )
        (progn
          (setq wb err-obj)
          (princ "\n✓ Workbook opened successfully")

          ; Get the active worksheet
          (setq ws (vlax-get-property wb 'ActiveSheet))
          (setq usedrange (vlax-get-property ws 'UsedRange))
          (setq rows (vlax-get-property usedrange 'Rows))
          (setq cols (vlax-get-property usedrange 'Columns))
          (setq row-count (vlax-get-property rows 'Count))
          (setq col-count (vlax-get-property cols 'Count))

          (if (< col-count 2)
            (progn
              (princ "\n*** ERROR: Worksheet must have at least 2 columns (tag, replacement)")
              (princ "\nPlease ensure column A contains tags and column B contains replacements")
              (setq error-occurred T)
            )
            (progn
              (princ (strcat "\n✓ Reading " (itoa row-count) " rows from worksheet..."))
              (setq row 1)
              (while (<= row row-count)
                ; Read cell values with error handling
                (setq cell-tag (vlax-get-property ws 'Cells row 1))
                (setq cell-repl (vlax-get-property ws 'Cells row 2))
                (setq tag (vlax-variant-value (vlax-get-property cell-tag 'Value)))
                (setq repl (vlax-variant-value (vlax-get-property cell-repl 'Value)))

                ; Convert variants to strings and handle nil/empty values
                (if (not tag) (setq tag ""))
                (if (not repl) (setq repl ""))
                (setq tag (vl-princ-to-string tag))
                (setq repl (vl-princ-to-string repl))
                (setq tag (trim tag))
                (setq repl (trim repl))

                (if (and (not (equal tag "")) (not (equal repl "")))
                  (setq pairs (cons (list tag repl) pairs))
                )
                (setq row (1+ row))
              )
              (princ (strcat "\n✓ Successfully loaded " (itoa (length pairs)) " valid pairs"))
            )
          )

          ; Clean up Excel COM objects properly
          (princ "\nClosing Excel...")
          (vlax-invoke wb 'Close :vlax-false)
          (vlax-invoke xl 'Quit)
          (vlax-release-object cols)
          (vlax-release-object rows)
          (vlax-release-object usedrange)
          (vlax-release-object ws)
          (vlax-release-object wb)
          (vlax-release-object xl)
          (princ "\n✓ Excel closed successfully")
        )
      )
    )
  )

  (if error-occurred
    '()
    (reverse pairs)
  )
)

(defun get-file-extension (filepath / dot-pos)
  (setq dot-pos (vl-string-search "." filepath :from-end T))
  (if dot-pos
    (strcase (substr filepath (+ dot-pos 2)))
    ""
  )
)

(defun read-pairs-from-file (file / ext)
  (setq ext (get-file-extension file))
  (cond
    ((or (equal ext "XLSX") (equal ext "XLS"))
     (read-excel-pairs file))
    ((equal ext "CSV")
     (read-csv-pairs file))
    (T
     (progn
       (princ "\nUnknown file format. Trying CSV format...")
       (read-csv-pairs file)
     ))
  )
)

(defun get-pt (obj / pt arr)
  (if (and obj (vlax-property-available-p obj 'InsertionPoint))
    (progn
      (setq arr (vlax-variant-value (vla-get-InsertionPoint obj)))
      (if (= (type arr) 'VECTOR)
        (list (car arr) (cadr arr) (caddr arr))
        (if (vlax-safearray-p arr)
          (vlax-safearray->list arr)
          nil
        )
      )
    )
    nil
  )
)

(defun get-text (obj / txt)
  (if (and obj (vlax-property-available-p obj 'TextString))
    (vla-get-TextString obj)
    ""
  )
)

(defun collect-text-objects (ms / list obj txt pt)
  (setq list '())
  (vlax-for ent ms
    (if (vlax-property-available-p ent 'TextString)
      (progn
        (setq txt (get-text ent))
        (setq pt  (get-pt ent))
        (if pt (setq list (cons (list ent txt pt) list)))
      )
    )
  )
  (reverse list)
)

(defun find-right-neighbor (tag-pt text-list / tx ty cand best dx dy best-item)
  (setq tx (car tag-pt) ty (cadr tag-pt) best nil best-item nil)
  (foreach item text-list
    (setq cand (nth 2 item)) ; point
    (setq dx (- (car cand) tx))
    (setq dy (- (cadr cand) ty))
    (if (and (> dx 0) (<= (abs dy) 1.0))
      (if (or (not best) (< dx (- (car best) tx)))
        (setq best cand best-item item)
      )
    )
  )
  best-item
)

(defun replace-text (obj new /)
  (if (and obj (vlax-property-available-p obj 'TextString))
    (vla-put-TextString obj new)
    nil
  )
)

(defun process-pair (tag repl txtObjs / total replacements entInfo entTxt entPt neighbor)
  (setq total 0 replacements 0)
  (foreach entInfo txtObjs
    (setq entTxt (nth 1 entInfo)
          entPt (nth 2 entInfo))
    (if (and entTxt (equal entTxt tag))
      (progn
        (setq total (1+ total))
        (setq neighbor (find-right-neighbor entPt txtObjs))
        (if neighbor
          (progn
            (replace-text (nth 0 neighbor) repl)
            (setq replacements (1+ replacements))
            (princ (strcat "\n  Replaced text for tag '" tag "' -> '" repl "'"))
          )
          (princ (strcat "\n  Warning: No right neighbor found for tag '" tag "'"))
        )
      )
    )
  )
  (list total replacements)
)

;; Main command: prompts for file and processes all tag/replacement pairs
(defun c:REPLTAG (/ filePath pairs acadObj doc ms txtObjs grandTotal grandRepl pair tag repl result pairTotal pairRepl)
  (princ "\n=== Replace Tags Tool ===")
  (princ "\nThis tool reads tag/replacement pairs from Excel or CSV files")
  (princ "\nand replaces text to the right of each tag in the drawing.")

  (setq filePath (getfiled "Select Excel or CSV file with tag,replacement pairs" "" "xlsx;xls;csv" 0))

  (if (not filePath)
    (progn
      (princ "\nNo file selected. Operation cancelled.")
      (princ)
    )
    (progn
      (princ (strcat "\nReading file: " filePath))
      (setq pairs (read-pairs-from-file filePath))

      (if (not pairs)
        (progn
          (princ "\n*** ERROR: No valid tag/replacement pairs found in file.")
          (princ "\nPlease ensure your file has two columns:")
          (princ "\n  Column 1: Tag text to find")
          (princ "\n  Column 2: Replacement text")
          (princ)
        )
        (progn
          (princ (strcat "\n✓ Loaded " (itoa (length pairs)) " tag/replacement pairs from file."))

          ; Initialize AutoCAD objects
          (setq acadObj (vlax-get-acad-object)
                doc     (vla-get-ActiveDocument acadObj)
                ms      (vla-get-ModelSpace doc)
                txtObjs (collect-text-objects ms)
                grandTotal 0
                grandRepl 0)

          (foreach pair pairs
            (setq tag (nth 0 pair)
                  repl (nth 1 pair)
                  result (process-pair tag repl txtObjs)
                  pairTotal (nth 0 result)
                  pairRepl (nth 1 result))

            (setq grandTotal (+ grandTotal pairTotal)
                  grandRepl (+ grandRepl pairRepl))

            (if (= pairTotal 0)
              (princ (strcat "\n  Note: Tag '" tag "' not found in drawing"))
            )
          )

          ; Print summary
          (princ "\n")
          (princ "\n=== SUMMARY ===")
          (princ (strcat "\nProcessed " (itoa (length pairs)) " tag/replacement pairs from file"))
          (princ (strcat "\nFound " (itoa grandTotal) " tag occurrences in drawing"))
          (princ (strcat "\nSuccessfully replaced " (itoa grandRepl) " text objects"))

          (if (< grandRepl grandTotal)
            (progn
              (princ (strcat "\n*** WARNING: "
                            (itoa (- grandTotal grandRepl))
                            " tags had no right neighbor for replacement"))
              (princ "\nConsider adjusting the vertical tolerance in find-right-neighbor if needed.")
            )
          )
          (princ "\n\n✓ Operation complete.")
          (princ)
        )
      )
    )
  )
  (princ)
) 

; EOF