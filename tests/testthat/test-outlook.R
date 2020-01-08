# Test that Outlook interface is working.
#
#

context('Outlook interface')




test_that("Outlook COM is created", {
   ol_app <- expect_silent(connect_outlook())

   expect_true(is_COM(ol_app))
   expect_true(is_outlook(ol_app))
   expect_equal(ol_app[['Name']], 'Outlook')

})


# Must ask for user action

test_that("Outlook COM is destroyed", {

   skip_if(!interactive())

   ol_app <- connect_outlook()

   expect_true(is_COM(ol_app))
   expect_true(is_outlook(ol_app))

   expect_silent(disconnect_outlook(ol_app))
})



# Message creation --------------------------------------------------------

test_that('A draft message is created silently.', {

   str <- 'Test draft message'

   ol_app <- connect_outlook()

   ol_msg <- expect_silent(create_draft(ol_app, body_plain = str, use_signature = FALSE, show_message = FALSE))

   # ol_msg <- create_draft(ol_app, body_plain = str, use_signature = FALSE, show_message = FALSE)

   expect_true(is_mail(ol_msg))

   expect_equal(stringr::str_trim(ol_msg[['Body']]), str)

   expect_silent(close_draft(ol_msg, save = FALSE))

   # disconnect_outlook(com)
})


test_that('A draft message is created with displaying.', {

   str <- 'Test draft message'

   ol_app <- connect_outlook()

   ol_msg <- expect_silent(create_draft(ol_app, body_plain = str, use_signature = FALSE, show_message = TRUE))
   expect_true(is_mail(ol_msg))

   expect_equal(stringr::str_trim(ol_msg[['Body']]), str)

   expect_silent(close_draft(ol_msg, save = FALSE))

   # disconnect_outlook(com)
})


