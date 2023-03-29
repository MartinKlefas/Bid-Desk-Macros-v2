# Bid Desk Macros

This is a set of tools intended to automate large parts of my work. While not widely useful some of the techniques and workarounds that I've cobbled together may be useful to other people.

The tools are essentially a suite of actions that read the clipboard and inbound emails to scrape data from each. In so doing they create and maintain records of my work, and allow the tool to forward incoming emails around the business as needed.

They have also evolved to have a server/client model possible, whereby a "server" can process rudimentary instructions sent by a client such as a remote worker who may not be able to be running the main server software on their device.

The backend code is written to work with a number of SQL options, though at present it needs to be refactored and code uncommented to switch between them. This is simply because the backend has changed over time with the security requirements of the IT team overseeing the work.

There is also legacy code for accessing and controlling a couple of vendor websites - but this is left in as an artifact - the controlled steps were eliminated by the vendors to increase productivity, but have a habit of coming back!
