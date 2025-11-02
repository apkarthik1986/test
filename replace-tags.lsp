<<<<<<< HEAD
; replace-tags.lsp
; AutoLISP script to replace text to the right of tags based on Excel or CSV file
; Usage: load this LISP in AutoCAD (appload or drag), then run REPLTAG command
; Input: Excel (.xlsx/.xls) or CSV file where column 1 = tag text, column 2 = replacement text
; Behavior: for each tag occurrence, finds the nearest text/mtxt/attribute to the right
; and replaces its content with the replacement value.
; Notes: Supports both Excel files (via COM) and CSV files (for portability).
; Excel COM requires full AutoCAD and appropriate system permissions.

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
      (if (> end (1- start)) (substr s start (- end start -1)) "")
    )
  )
)

(defun read-csv-pairs (file / fh line parts tag repl pairs)
  (setq pairs '())
  (if (setq fh (open file "r"))
    (progn
      (while (setq line (read-line fh))
        (if (not (equal (trim line) ""))
          (progn
            (setq parts (split line ","))
            (when (>= (length parts) 2)
              (setq tag (trim (nth 0 parts)))
              (setq repl (trim (nth 1 parts)))
              (if (and (not (equal tag "")) repl)
                (setq pairs (cons (list tag repl) pairs))
              )
            )
          )
        )
      )
      (close fh)
    )
    (progn (princ (strcat "\nCould not open CSV file: " file)) )
  )
  (reverse pairs)
)

(defun read-excel-pairs (file / xl wb ws pairs row-count col-count row tag repl error-occurred)
  (setq pairs '() error-occurred nil)
  (if (vl-catch-all-error-p
        (vl-catch-all-apply 'vlax-get-or-create-object '("Excel.Application")))
    (progn
      (princ "\nExcel COM not available. Please use CSV format instead.")
      (setq error-occurred T)
    )
    (progn
      (princ "\nOpening Excel file via COM...")
      (setq xl (vlax-get-or-create-object "Excel.Application"))
      (if (vl-catch-all-error-p
            (setq wb (vl-catch-all-apply 'vla-open 
                       (list (vlax-get-property xl 'Workbooks) file))))
        (progn
          (princ (strcat "\nCould not open Excel file: " file))
          (if xl (vlax-release-object xl))
          (setq error-occurred T)
        )
        (progn
          (setq ws (vlax-get-property wb 'ActiveSheet))
          (setq row-count (vlax-get-property (vlax-get-property ws 'UsedRange) 'Rows 'Count))
          (setq col-count (vlax-get-property (vlax-get-property ws 'UsedRange) 'Columns 'Count))
          
          (if (< col-count 2)
            (progn
              (princ "\nWorksheet must have at least 2 columns (tag, replacement)")
              (setq error-occurred T)
            )
            (progn
              (princ (strcat "\nReading " (itoa row-count) " rows from worksheet..."))
              (setq row 1)
              (while (<= row row-count)
                (setq tag (vlax-variant-value 
                            (vlax-get-property ws 'Cells row 1 'Value)))
                (setq repl (vlax-variant-value 
                             (vlax-get-property ws 'Cells row 2 'Value)))
                
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
            )
          )
          
          ; Clean up Excel COM objects
          (vlax-invoke wb 'Close :vlax-false)
          (vlax-invoke xl 'Quit)
          (vlax-release-object ws)
          (vlax-release-object wb)
          (vlax-release-object xl)
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

(defun find-right-neighbor (tag-pt text-list / tx ty cand best dx dy)
  (setq tx (car tag-pt) ty (cadr tag-pt) best nil)
  (foreach item text-list
    (setq cand (nth 2 item)) ; point
    (if (and cand
             (> (car cand) tx) ; to the right
             (<= (abs (- (cadr cand) ty)) 1.0) ; vertical tolerance (units)
        )
      (progn
        (setq dx (- (car cand) tx))
        (if (or (null best) (< dx (car best)))
          (setq best (list dx item))
        )
      )
    )
  )
  (if best (cadr best) nil)
)

(defun replace-text (target-obj newstr / )
  (if (and target-obj (vlax-property-available-p target-obj 'TextString))
    (vla-put-TextString target-obj newstr)
  )
)

(defun c:REPLTAG (/ filePath pairs acadObj doc ms txtObjs total replacements)
  (princ "\nReplace Tags: Supports Excel (.xlsx/.xls) and CSV files.")
  (setq filePath (getfiled "Select Excel or CSV file with tag,replacement" "" "xlsx;xls;csv" 0))
  (if (not filePath)
    (progn (princ "\nNo file selected. Aborting.") (princ))
    (progn
      (setq pairs (read-pairs-from-file filePath))
      (if (not pairs)
        (princ "\nNo pairs found in file or file could not be read. Aborting.")
        (progn
          (princ (strcat "\nLoaded " (itoa (length pairs)) " tag/replacement pairs."))
          (setq acadObj (vlax-get-acad-object)
                doc     (vla-get-ActiveDocument acadObj)
                ms      (vla-get-ModelSpace doc))
          (setq txtObjs (collect-text-objects ms))
          (princ (strcat "\nFound " (itoa (length txtObjs)) " text objects in ModelSpace."))
          (setq total 0 replacements 0)
          (foreach pr pairs
            (setq tag (nth 0 pr) repl (nth 1 pr))
            (foreach entInfo txtObjs
              (setq entObj (nth 0 entInfo) entTxt (nth 1 entInfo) entPt (nth 2 entInfo))
              (if (and entTxt (equal entTxt tag))
                (progn
                  (setq total (1+ total))
                  (setq neighbor (find-right-neighbor entPt txtObjs))
                  (if neighbor
                    (progn
                      (replace-text (nth 0 neighbor) repl)
                      (setq replacements (1+ replacements))
                      (princ (strcat "\nReplaced text for tag: " tag " -> " repl))
                    )
                    (princ (strcat "\nNo right neighbor found for tag: " tag))
                  )
                )
              )
            )
          )
          (princ (strcat "\nSUMMARY: Processed " (itoa total) " tag occurrences; replaced " (itoa replacements) " texts."))
        )
      )
    )
  )
  (princ)
)

; EOF
=======
; replace-tags.lsp
; AutoLISP script to replace text to the right of tags based on Excel or CSV file
; Usage: load this LISP in AutoCAD (appload or drag), then run REPLTAG command
; Input: Excel (.xlsx/.xls) or CSV file where column 1 = tag text, column 2 = replacement text
; Behavior: for each tag occurrence, finds the nearest text/mtxt/attribute to the right
; and replaces its content with the replacement value.
; Notes: Supports both Excel files (via COM) and CSV files (for portability).
; Excel COM requires full AutoCAD and appropriate system permissions.

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
      (if (> end (1- start)) (substr s start (- end start -1)) "")
    )
  )
)

(defun read-csv-pairs (file / fh line parts tag repl pairs)
  (setq pairs '())
  (if (setq fh (open file "r"))
    (progn
      (while (setq line (read-line fh))
        (if (not (equal (trim line) ""))
          (progn
            (setq parts (split line ","))
            (when (>= (length parts) 2)
              (setq tag (trim (nth 0 parts)))
              (setq repl (trim (nth 1 parts)))
              (if (and (not (equal tag "")) repl)
                (setq pairs (cons (list tag repl) pairs))
              )
            )
          )
        )
      )
      (close fh)
    )
    (progn (princ (strcat "\nCould not open CSV file: " file)) )
  )
  (reverse pairs)
)

(defun read-excel-pairs (file / xl wb ws pairs row-count col-count row tag repl error-occurred)
  (setq pairs '() error-occurred nil)
  (if (vl-catch-all-error-p
        (vl-catch-all-apply 'vlax-get-or-create-object '("Excel.Application")))
    (progn
      (princ "\nExcel COM not available. Please use CSV format instead.")
      (setq error-occurred T)
    )
    (progn
      (princ "\nOpening Excel file via COM...")
      (setq xl (vlax-get-or-create-object "Excel.Application"))
      (if (vl-catch-all-error-p
            (setq wb (vl-catch-all-apply 'vla-open 
                       (list (vlax-get-property xl 'Workbooks) file))))
        (progn
          (princ (strcat "\nCould not open Excel file: " file))
          (if xl (vlax-release-object xl))
          (setq error-occurred T)
        )
        (progn
          (setq ws (vlax-get-property wb 'ActiveSheet))
          (setq row-count (vlax-get-property (vlax-get-property ws 'UsedRange) 'Rows 'Count))
          (setq col-count (vlax-get-property (vlax-get-property ws 'UsedRange) 'Columns 'Count))
          
          (if (< col-count 2)
            (progn
              (princ "\nWorksheet must have at least 2 columns (tag, replacement)")
              (setq error-occurred T)
            )
            (progn
              (princ (strcat "\nReading " (itoa row-count) " rows from worksheet..."))
              (setq row 1)
              (while (<= row row-count)
                (setq tag (vlax-variant-value 
                            (vlax-get-property ws 'Cells row 1 'Value)))
                (setq repl (vlax-variant-value 
                             (vlax-get-property ws 'Cells row 2 'Value)))
                
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
            )
          )
          
          ; Clean up Excel COM objects
          (vlax-invoke wb 'Close :vlax-false)
          (vlax-invoke xl 'Quit)
          (vlax-release-object ws)
          (vlax-release-object wb)
          (vlax-release-object xl)
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

(defun find-right-neighbor (tag-pt text-list / tx ty cand best dx dy)
  (setq tx (car tag-pt) ty (cadr tag-pt) best nil)
  (foreach item text-list
    (setq cand (nth 2 item)) ; point
    (if (and cand
             (> (car cand) tx) ; to the right
             (<= (abs (- (cadr cand) ty)) 1.0) ; vertical tolerance (units)
        )
      (progn
        (setq dx (- (car cand) tx))
        (if (or (null best) (< dx (car best)))
          (setq best (list dx item))
        )
      )
    )
  )
  (if best (cadr best) nil)
)

(defun replace-text (target-obj newstr / )
  (if (and target-obj (vlax-property-available-p target-obj 'TextString))
    (vla-put-TextString target-obj newstr)
  )
)

(defun c:REPLTAG (/ filePath pairs acadObj doc ms txtObjs total replacements)
  (princ "\nReplace Tags: Supports Excel (.xlsx/.xls) and CSV files.")
  (setq filePath (getfiled "Select Excel or CSV file with tag,replacement" "" "xlsx;xls;csv" 0))
  (if (not filePath)
    (progn (princ "\nNo file selected. Aborting.") (princ))
    (progn
      (setq pairs (read-pairs-from-file filePath))
      (if (not pairs)
        (princ "\nNo pairs found in file or file could not be read. Aborting.")
        (progn
          (princ (strcat "\nLoaded " (itoa (length pairs)) " tag/replacement pairs."))
          (setq acadObj (vlax-get-acad-object)
                doc     (vla-get-ActiveDocument acadObj)
                ms      (vla-get-ModelSpace doc))
          (setq txtObjs (collect-text-objects ms))
          (princ (strcat "\nFound " (itoa (length txtObjs)) " text objects in ModelSpace."))
          (setq total 0 replacements 0)
          (foreach pr pairs
            (setq tag (nth 0 pr) repl (nth 1 pr))
            (foreach entInfo txtObjs
              (setq entObj (nth 0 entInfo) entTxt (nth 1 entInfo) entPt (nth 2 entInfo))
              (if (and entTxt (equal entTxt tag))
                (progn
                  (setq total (1+ total))
                  (setq neighbor (find-right-neighbor entPt txtObjs))
                  (if neighbor
                    (progn
                      (replace-text (nth 0 neighbor) repl)
                      (setq replacements (1+ replacements))
                      (princ (strcat "\nReplaced text for tag: " tag " -> " repl))
                    )
                    (princ (strcat "\nNo right neighbor found for tag: " tag))
                  )
                )
              )
            )
          )
          (princ (strcat "\nSUMMARY: Processed " (itoa total) " tag occurrences; replaced " (itoa replacements) " texts."))
        )
      )
    )
  )
  (princ)
)

; EOF
>>>>>>> b8b59636fc143418f0c3514f3d0a474a365ea994
