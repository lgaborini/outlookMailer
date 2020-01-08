# Connect to Outlook, then create an email message.
#
#


#' Connect to Outlook and create an e-mail message.
#'
#' Connect to Outlook and create an e-mail message.
#'
#' @param ol_app An object that represents an Outlook application instance (class: class `COMIDispatch.`)
#' @param addr_from optional fields to fill
#' @param addr_to optional fields to fill
#' @param addr_cc optional fields to fill
#' @param subject optional fields to fill
#' @param html_body HTML contents of the message. Takes precedence over body
#' @param body plain text contents of the message
#' @param attachments path to file(s) to attach
#' @param use_signature if TRUE, get the signature (if automatically added) from the blank message, than put it back after the message is created.
#' @param show_message if TRUE, show the message for editing after creation
#'
#' @return a COM object that binds to an e-mail window ('Outlook.MailItem')
#' @export
#' @import RDCOMClient
#' @importFrom fs file_exists
#' @importFrom glue glue
#' @importFrom rlang abort
#' @examples
#' \dontrun{
#'
#' com <- connect_outlook()
#' msg <- create_draft(com, addr_to = "foo@bar.com")
#'
#' # Attachments
#' msg <- create_draft(con, attachments = c('foo.txt', 'foo2.txt'))
#'
#' }
create_draft <- function(ol_app,
                         addr_from = NULL,
                         addr_to = NULL,
                         addr_cc = NULL,
                         subject = NULL,
                         body_html = NULL,
                         body_plain = NULL,
                         attachments = NULL,
                         use_signature = TRUE,
                         show_message = TRUE) {

   if (!require('RDCOMClient')) {
      stop('Please install missing "RDCOMClient" package.\n> devtools::install_github("omegahat/RDCOMClient")')
   }

   stopifnot(is_outlook(ol_app))

   # Code from:
   #
   # https://stackoverflow.com/questions/42972222/how-to-send-mails-from-outlook-using-r-rdcomclient-using-latest-version
   #

   ol_mail <- ol_app$CreateItem(0)
   stopifnot(is_mail(ol_mail))

   # Grab the signature from the blank message, then paste it back
   #
   # It works only if the signature is automatically added.
   if (use_signature) {
      signature <- ol_mail[["HTMLBody"]]
   }
#
#    addr_to <- 'test@test.com'
#    addr_cc <- ''
#    subject <- 'Fake email'
#    body <- 'Plain text.'

   ## configure  email parameter
   if (!is.null(addr_from)) ol_mail[["From"]]<- addr_from
   if (!is.null(addr_to)) ol_mail[["To"]]<- addr_to
   if (!is.null(addr_cc)) ol_mail[["CC"]] <- addr_cc
   if (!is.null(subject)) ol_mail[["subject"]] <- subject

   if (!is.null(body_html)) {
      ol_mail[["HTMLBody"]] <- body_plain

      if (!is.null(body_plain)) {
         message('Supplied HTML body, discarding supplied plain text body.')
      }

   } else {
      if (!is.null(body_plain)) {
         ol_mail[["Body"]] <- body_plain
      }
   }

   if (!is.null(attachments)) {

      # Check that files exist
      stopifnot(is.character(attachments))
      for (f in attachments) {
         if (!fs::file_exists(f)) {
            rlang::abort(glue::glue('attachments: file {f} does not exist.'), class = 'attachment_not_found')
         }
         ol_mail[['Attachments']]$Add(f)
      }
   }

   # Paste back the signature
   if (use_signature) {
      html_body_final <- ol_mail[["HTMLBody"]]
      ol_mail[["HTMLBody"]] <- paste0(html_body_final, '<p>', signature, '</p>')
   }

   # Show the message
   if (show_message){

      stopifnot(has_COM_method(ol_mail, 'Display'))

      Sys.sleep(0.5)
      ol_mail$Display()
   }

   ol_mail
   # ol_app$Quit()

}



#' Close an Outlook message window.
#'
#' @param ol_mail a COM object that binds to an e-mail window ('Outlook.MailItem')
#' @param save if TRUE, save to drafts, else discard without confirmation
#' @export
close_draft <- function(ol_mail, save = FALSE) {

   stopifnot(is_mail(ol_mail))

   # ol_mail$Close(2) # olPromptForSave

   if (save) {
      ol_mail$Close(0) # olSave
   } else {
      ol_mail$Close(1) # olDiscard
   }

   invisible(NULL)
}

