;; gui.lsp - Simple Ltk GUI Example
;; Make sure Ltk is installed: https://github.com/tfeb/ltk

(require :ltk)

(defun main ()
  (ltk:with-ltk ()
    (let ((label (ltk:make-widget 'ltk:label :text "Hello, Lisp GUI!")))
      (ltk:pack label)
      (ltk:make-widget 'ltk:button :text "Click Me"
                       :command (lambda ()
                                  (ltk:configure label :text "Button Clicked!")))
      (ltk:pack ltk:*last-widget*))))

(main)
