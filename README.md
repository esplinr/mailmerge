mailmerge
=========

Simple script for sending an email template to everyone in a CSV, with variable substitution.

This works great for organizing a conference. Please don't use it to send
anything stupid.

Notes
-------
* Script does not handle well UTF8 in the To: field.
* Gmail has a limit of 100 emails sent in a 90 second period. Then it closes the
  socket and you have to wait 90 seconds before sending again. It might be an
  exponential back-off, as while testing the script the time kept getting longer.
