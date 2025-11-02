;;; simple-gui.el --- Simple Emacs Lisp GUI Example

(defun simple-gui-show-message ()
  "Show a message when the button is clicked."
  (message "Button clicked!"))

(defun simple-gui-open ()
  "Open a buffer with a simple button."
  (interactive)
  (let ((buf (get-buffer-create "*Simple GUI*")))
    (with-current-buffer buf
      (erase-buffer)
      (insert "Welcome to Simple Emacs GUI!\n\n")
      (insert-text-button "Click Me"
                         'action (lambda (_) (simple-gui-show-message))
                         'follow-link t)
      (insert "\n\nClick the button above to see a message in the minibuffer.")
      (goto-char (point-min)))
    (pop-to-buffer buf)))

;; To use: M-x load-file RET simple-gui.el RET
;; Then run: M-x simple-gui-open

(provide 'simple-gui)
