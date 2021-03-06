
<!-- README.md is generated from README.Rmd. Please edit that file -->

# outlookMailer

<!-- badges: start -->

<!-- badges: end -->

The goal of outlookMailer is to create an R interface between Microsoft
Outlook and R, to compose messages from R using Outlook’s native
window.  
It works only on Windows.

## Installation

### Requirements

The package requires access to Windows DCOM objects, provided by
[RDCOMClient](http://www.omegahat.net/RDCOMClient/).  
It is no longer on CRAN, but can be installed from
[GitHub](https://github.com/omegahat/RDCOMClient):

``` r
remotes::install_github('omegahat/RDCOMClient')
```

### The package

The package is not on CRAN, install the developer version from here:

``` r
remotes::install_github('lgaborini/outlookMailer')
```

## Manual

[Here](https://lgaborini.github.io/outlookMailer/) (built with
[pkgdown](https://pkgdown.r-lib.org/)📦)

## Example

The package wraps the COM interface to Microsoft Outlook with
user-friendly R functions.

One retrieves/sets properties with `[[` and `]]` accessors.  
One calls methods with the `$` accessor.

``` r
library(outlookMailer)

# Create a connection to Outlook
con <- connect_outlook()

# Create a message and show it
msg <- create_draft(con, 
                    addr_to = 'foo@bar.com', 
                    body_plain = 'Body of the message', 
                    use_signature = TRUE,
                    show_message = FALSE)

# Optionally modify properties
msg[['Subject']] <- 'Subject of the message'

# Show the message
msg$Display()

# Send the message (caution!)
# msg$Send()

# Close and save to drafts
close_draft(msg, save = TRUE)
```

Most functions are pipe-friendly:

``` r
connect_outlook() %>% 
  create_draft(body_plain = 'Test') %>% 
  close_draft(save = TRUE) %>%
  disconnect_outlook()
```

### Attachments

Multiple attachments can be supplied as a vector of paths.

``` r
msg <- create_draft(con, attachments = c('foo.txt', 'foo2.txt'))
```

### Messages saved as files

Messages saved as `.msg` or `.eml` can be opened and modified:

``` r
con <- connect_outlook()

msg <- open_msg(con, path_msg = 'sample.msg', show_message = TRUE)
```

### COM utilities

The package provides a function check if a COM object has a property:

``` r
has_COM_property(con, 'Name')
con[['Name']]
# 'Outlook'
```

## TODO

  - [ ] Fix the signature (HTML/plain text conversion)
  - [ ] Save and discard messages
  - [ ] Make it pipe-friendly (e.g. `set_from`, `set_body`, …)
  - [ ] Improve robustness to Windows calls
  - [ ] Re-implement `has_COM_method`
  - [ ] Move to S3 objects
