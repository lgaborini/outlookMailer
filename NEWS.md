# outlookMailer 0.1.3

* `open_msg` no longer lock files, it copies the message to temporary directory. Also, it works on local drives.

# outlookMailer 0.1.2

* Removed `has_COM_method`: not 100% reliable.
* More README docs.

# outlookMailer 0.1.1

* Bugfix: `Outlook.MailItem` does not have any `From` property (it is called `Sender` instead).
* Add explicit `disconnect_outlook` in tests.
* Bugfix: signature is not correctly pasted in HTML e-mails. Do not paste on HTML e-mails for now.

# outlookMailer 0.1

* Add `open_msg` function to open messages from files.

# outlookMailer 0.0.1

* Added a `NEWS.md` file to track changes to the package.
