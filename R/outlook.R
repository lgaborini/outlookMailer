# COM interface to Outlook
#
#


#' Create a COM Outlook.Application instance.
#'
#' @return An object of class `COMIDispatch.` that represents an Outlook application instance.
#' @export
#'
#' @examples
#' \dontrun{
#' com <- connect_outlook()
#' }
connect_outlook <- function(wait_seconds = 3) {

   if (!require('RDCOMClient', quietly = TRUE)) {
      stop('Please install missing "RDCOMClient" package.\n> devtools::install_github("omegahat/RDCOMClient")')
   }

   # Code from:
   #
   # https://stackoverflow.com/questions/42972222/how-to-send-mails-from-outlook-using-r-rdcomclient-using-latest-version
   #

   ol_app <- RDCOMClient::COMCreate('outlook.application')

   wait_start <- Sys.time()
   while (as.numeric(Sys.time() - wait_start) < wait_seconds) {

      if (is_outlook(ol_app)) {
         # message(sprintf('Created! (took %f seconds)', as.numeric(Sys.time() - wait_start)))
         # break
         return(ol_app)
      }
   }

   if (!is_outlook(ol_app)) {
      stop('cannot spawn Outlook window.')
   }

   # Wait for Outlook spawn
   # Sys.sleep(wait_seconds)

   ol_app
}

#' Close an Outlook COM instance.
#'
#' @param com An object of class `COMIDispatch.` that represents an Outlook application instance.
#' @return nothing
#' @export
#' @examples
#' \dontrun{
#' disconnect_outlook(com)
#' }
disconnect_outlook <- function(com) {
   if (is_outlook(com)) {
      com$Quit()
   }
   invisible(NULL)
}



# Is-methods --------------------------------------------------------------



#' Check if an object is a COM object.
#'
#' Check if an object is a COM object.
#'
#' @param x any object
#' @return TRUE or FALSE
#' @keywords internal
is_COM <- function(x) {
   is(x, "COMIDispatch")
}



#' Check if an object is an Outlook application.
#'
#' Check if an object is an Outlook application.
#' TRUE if the object is a COM object with the property `Name`, with value `Outlook`.
#'
#' @return TRUE if the object is a binding to Outlook Application instance.
#' @export
#' @inheritParams is_COM
is_outlook <- function(x) {

   if (!is_COM(x)) return(FALSE)

   has_name <- has_COM_property(x, 'Name')
   if (!has_name) return(FALSE)

   name_value <- x[['Name']]

   if (name_value != 'Outlook') return(FALSE)


   # has_method <- has_COM_method(x, 'CreateItem')

   return(TRUE)
   # return(has_method)
}




#' Check if an object is an Outlook message.
#'
#' Check if an object is an Outlook message.
#'
#' @return TRUE if the object is a binding to Outlook.MailItem instance.
#' @export
#' @inheritParams is_outlook
is_mail <- function(x) {

   if (!is_COM(x)) return(FALSE)
   if (!has_COM_property(x, 'Class')) return(FALSE)
   if (!has_COM_property(x, 'To')) return(FALSE)
   if (!has_COM_property(x, 'CC')) return(FALSE)
   if (!has_COM_property(x, 'Subject')) return(FALSE)
   # if (!has_COM_method(x, 'Display')) return(FALSE)

   ret <- (x[['Class']] == 43)           # olMailItem
   ret
}


# Has-methods -------------------------------------------------------------


#' Check if a COM object has a property (attribute)
#'
#' @param prop_name a single property name to query, as a character
#' @return TRUE if x is a COM object with the given property
#' @export
#' @importFrom rlang catch_cnd exec
#' @inheritParams is_COM
has_COM_property <- function(x, prop_name) {

   stopifnot(is.character(prop_name))
   stopifnot(length(prop_name) == 1)

   if (!is_COM(x)) return(FALSE)

   cnd <- rlang::catch_cnd(x[[prop_name]])

   # No condition is thrown
   if (is.null(cnd)) {
      return(TRUE)
   }

   if (stringr::str_detect(cnd$message, '^Cannot locate [0-9]+ name\\(s\\)')) {
      return(FALSE)
   }
   # Throw any other condition
   stop(cnd)
}




