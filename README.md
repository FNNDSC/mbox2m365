# Abstract

This repository contains a simple python script that (1) processes an
`mbox` style mailbox file, (2) parses out (typically) the last message
and then (3) re-transmits this using the CLI `m365` toolset. While this
script can be used in a completely stand-alone manner to extract a
message from an `mbox` and then *send* using `m365`, it’s main utility
is when used in a full ecosystem including an email client and `postfix`
server.

The CLI `m365` offers the ability to use Outlook/Office365 to send
emails. However, this tool is not suited to use as a mail client — it
merely offers a means to transmit a message from the command line. More
useful would be if existing mail clients could leverage `m365` natively.
Currently, however, this is not available. Therefore, `mbox2m365` was
created to complete a "missing link" in a full email experience. Using
this tool in a configured environment, it is now possible to use *any*
email client to transmit messages using Outlook/Office365.

# TL;DR

Even if you’re impatient, it’s a good idea to read this document.
However, assuming you’ve already built the infrastructure, assuming you
have installed `m365` (and logged into your Institution from `m365`),
have assuming you have setup your email client to email to a dummy local
postfix account, then you can install `mbox2m365` with a simple

``` bash
pip install mbox2m365
```

note at time of writing the PyPI install might not be available. In that
case, simply clone this repo and do a

``` bash
# From the repo root dir
pip install -U .
```

and fire it off by setting up the monitor on the `mbox` file (here we
assume the mbox belongs to `rudolph`)

``` bash
cd /var/mail
find . | entr mbox2m365 --mbox rudolph --b64_encode         \
                        --sendFromFile --cleanUp            \
                        --waitForStragglers 5
```

(Clearly the above relies on [`entr`](https://github.com/eradman/entr).
Consult your distro’s repos for the appropriate install. In arch, this
is `yay -S entr`, in Ubuntu, this is `sudo apt install entr`)

# Security alert!

Wait! **Danger, Will Robinson**! Doesn’t this violate security policies?
I was told my Institution does *not* allow clients like thunderbird to
connect to the Outlook server to send emails!

Well, the devil is in the details as the saying goes. This solution does
*not* violate policy since the "unauthenticated" email client is *not*
ever connecting/authenticating to the Outlook server. Rather, the
"unauthenticated" email client sends messages to a specially configured
server that "saves" the emails to a file. Then, a separate program
teases out the new arrivals and uses the **authenticated** `m365`
program to actually talk to and leverage Outlook to send the email. This
*separate program* that bridges the mail file to `m365` is none other
than this python module.

So to be pendantic: all credentialing and authentication is handled by
`m365` and *not* this tool. The end user is still required to
login/authenticate using `m365`. If this has not happened, then
`mbox2m365` will not work.

Again, this repo contains **no** authentication tokens/passwords/etc.

# Limitations

Let’s get these out of the way first. This is *not* a full solution for
linking a third party email client to Outlook. It gets *close*, IMHO.
The limitations are:

- This is not a multi-user solution. All email transmitted to the
  intermediate helper server is added to a single user `mbox` which when
  processed by this script means **all** email in that `mbox` will
  appear to have been transmitted by that user. While in theory it
  should be possible to support multiple users, either by having a
  separate helper server per user or by having multiple local users on
  the mail server and separate `mbox` files, the current solution is
  clearly not scalable. It is decent though for the single user case.

- Fundmanentally, this solution is limited first and foremost by the
  capabilities of the CLI `m365 outlook mail send` functionality. Recent
  updates to `m365` have expanded functionality of this bridge — most
  notably attachments are now supported. Still pending is full `bcc`
  support — currently all `bcc` recipients are switched to `cc`.

Other limitations stem from the fact the bridge needs to process the
`mbox` file *de novo* each time mail is received by the helper server.
This means that time to process will increase linearly with `mbox` size.
The real world implications of this are still being explored.

Finally, the bridge does make a best effort attempt to *wait* until the
`mbox` has stabilized before analyzing it. When sending emails to
multiple recipients, *each* recipient gets a single complete copy in the
`mbox` with a single address (the multiple recipient addresses are not
conserved in the `mbox` but *are* processed by this bridge). While
arguably wasteful, the implication is that while all these recipient
copies are appended to the `mbox`, the bridge should wait until all have
been added before processing. Currently the bridge attempts this by
examining the `mbox` file size when called, and waiting a small delta
time interval and checking if the size has changed again. It will wait
until the size is stable before continuing. The TL;DR is this could
result in a processing delay.

# Introduction

The migration/adoption of Office365 by institutions often poses issues
and problems for users wishing to use different tools for email
transmission, particularly on platforms such as Linux.

While it is possible to access various Office365 applications,
particularly Outlook, via a web interface, this is often not sufficient
for several classes of users, particularly those who find the web
interface cumbersome (for instance, the web interface has no (or
limited) support for keyboard shortcuts, it is cumbersome if not
impossible to automate tasks involving event-driven email) and/or for
users who have an existing and efficient email workflow using clients
such as `thunderbird` or `mutt`.

This small repo houses a small and simple python application with some
supporting documentation that provides a solution to the problem of
sending email from an instituion’s managed domain by leveraging
`postfix` and some `mbox` processing and then using the CLI tool `m365`.

# Method Summary

The solution requires some seemingly contortuous steps, but in reality
is rather simple and can be summarized as follows:

## Reading email

Reading emails that are locked away in an Outlook server is best
effected by simply adding a forward incoming email rule from your
Institution Outlook Server (IOS) to an externally accessible email
provider (such as `gmail`), allowing this email to be read easily by
tools such as `thunderbird` or `mutt`.

    ┌─────────────────────────────────┐
    │IOS that receives incoming email │
    └┬────────────────────────────────┘
     │
     └─────┐
           │
          ┌O─────────────────────┐
          │forwardRule(<message>)│
          └┬─────────────────────┘
     ┌─────┘
     │
    ┌O────┐
    │gmail│
    └┬────┘
     └─────┐
           │
          ┌O──────────────────────┐
          │clientAccess(<message>)│
          └O──────────────────────┘
     ┌─────┘
     │
    ┌┴───────────────────────┐
    │thunderbird / mutt / etc│
    └────────────────────────┘

## Sending email

The message is now outside of Outlook, and if the Institution does not
allow non-authorized clients (often this means they only allow Microsoft
tools) to connect to the Outlook server, the following work around will
help. Essentially, the outside client should be configured to send email
using a properly setup `postfix` server that simply copies the target
email to an `mbox` file.

This `mbox` file is then monitored for any changes, and on a change
(assumed to mean a new email message has been appended), a new process
is fired off to parse off the latest message and then use the command
line `m365` CLI tool to have the IOS send the email.

    ┌───────────────────────┐
    │thundebird / mutt /etc │
    └┬──────────────────────┘
     │
     └─────┐
           │
          ┌O────────────────────────────┐
          │sendmail(<message>) (postfix)│
          └┬────────────────────────────┘
     ┌─────┘
     │
    ┌O────┐
    │mbox │
    └┬────┘
     └─────┐
           │
          ┌O────────┐
          │mbox2m365│ <--- this repo!
          └┬────────┘
     ┌─────┘
     │
    ┌O─────────────┐
    │m365 <message>│
    └┬─────────────┘
     │
    ┌O────────────────────────────────┐
    │IOS that transmits outgoing email│
    └─────────────────────────────────┘

# `mbox2m365`

While all the building blocks to effect the solution exist, the one
missing piece is the `mbox` to `m365` block, which is provided for by
this rather simple python script.

# Setup your helper mail server, `postfix`

First, install `postfix`

## Arch

``` bash
yay -S postfix
```

## Ubuntu

``` bash
sudo apt install postfix
```

# `transport`

Now, edit the `transport` file.

``` bash
sudo bash
cd /etc/postfix
cp transport transport.orig
echo "* local:rudolph" >> transport
```

# `main.cf`

For the `main.cf` file, do

``` bash
# Assuming you are still in the /etc/postfix dir in a sudo bash shell...
cp main.cf main.cf.orig
echo "mydomain = pangea.net" >> main.cf
echo "luser_relay = rudolph@pangea.net"
echo "transport_maps = hash:/etc/postfix/transport" >> main.cf
```

# enable/restart the services

``` bash
sudo systemctl enable postfix.service
sudo systemctl restart postfix.service
```

# Email client

Simply configure your email client to use the machine running `postfix`
as your email server. All emails will be appended to the `transport`
user’s `mbox` file.

# Fire up `mbox2m365`

The final piece of the puzzle:

``` bash
cd /var/mail
find . | entr mbox2m365 --mbox rudolph --b64_encode         \
                        --sendFromFile --cleanUp            \
                        --waitForStragglers 5
```

*-30-*
