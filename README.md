# els-tag-repair

Server-side injection of tag or warning message strings directly in the email body is not only a technically unfortunate implementation but has numerous drawbacks for users (e.g. it is not automatically stripped from replies)

This Outlook VBA macro repairs the email body by removing such strings from from the email body and sets a more appropriatecustom category instead.

This macro achieves the following:
- it removes a specific string from the email body of any incoming email
- and sets a custom category instead so emails are still marked
- the string is currently set to "External email: use caution" with surrounding asterisks

## Screenshots

![Without macro](https://raw.githubusercontent.com/jerogee/els-tag-repair/master/img/ss_without.png)

![With macro](https://raw.githubusercontent.com/jerogee/els-tag-repair/master/img/ss_with.png)


## Installation

The easiest way to install this macro is:
* Press `Alt` + `F11` to bring up the `VBA` environment.
* In the Project pane, under `Project1`, double-click the built-in `ThisOutlookSession` module to open it.
* Copy & paste the macro code (use Github's raw view) into it.
* Close the editor.
* To be able to run the macro: In Outlook 2007 and higher, the macro security settings are in the `Tools` | `Trust Center` dialog. Set macro security to `Warn on all macros`. Restart Outlook.

