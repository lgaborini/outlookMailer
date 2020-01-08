
<!-- README.md is generated from README.Rmd. Please edit that file -->

# outlookMailer

<!-- badges: start -->

<!-- badges: end -->

The goal of outlookMailer is to create an R interface between Microsoft
Outlook and R, to compose messages from R using Outlook’s native window.

## Installation

### Requirements

``` r
remotes::install_github('omegahat/RDCOMClient')
#> Skipping install of 'RDCOMClient' from a github remote, the SHA1 (6e1e5806) has not changed since last install.
#>   Use `force = TRUE` to force installation
```

### The package

You can install the released version of outlookMailer from
[CRAN](https://CRAN.R-project.org) with:

``` r
install.packages("outlookMailer")
```

## Example

The package wraps the COM interface to Microsoft Outlook with
user-friendly R functions.

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

### Attachments

Multiple attachments can be supplied as a vector of paths.

``` r
msg <- create_draft(con, attachments = c('foo.txt', 'foo2.txt'))
```
