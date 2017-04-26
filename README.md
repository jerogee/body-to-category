# els-tag-repair

Server-side injection of tag or warning message strings directly in the email body is not only a technically unfortunate implementation but has numerous drawbacks for users (e.g. it is not automatically stripped from replies)

This Outlook VBA macro repairs the email body by removing such strings from from the email body and sets a more appropriate [custom category](https://support.office.com/en-us/article/Create-and-assign-color-categories-a1fde97e-15e1-4179-a1a0-8a91ef89b8dc) instead.

This macro achieves the following:
- it removes a specific string from the email body of any incoming email
- and sets a custom category instead so emails are still marked
- the string is currently set to "External email: use caution" with surrounding asterisks (line 23)

## Screenshots

Without macro:
![Without macro](https://raw.githubusercontent.com/jerogee/els-tag-repair/master/img/ss_without.png)

With macro:
![With macro](https://raw.githubusercontent.com/jerogee/els-tag-repair/master/img/ss_with.png)


## Installation

The easiest way to install this macro is:
* Press `Alt` + `F11` to bring up the `VBA` environment.
* In the Project pane, under `Project1`, double-click the built-in `ThisOutlookSession` module to open it.
* Copy & paste the macro code from `tagFromBodyToCategory.vba` (use Github's raw view) into it.
* Close the editor.
* To be able to run the macro, set the security settings appropriately: In Outlook 2007 and higher, the macro security settings are in `Options` | `Trust Center` | `Trust Center Settings...` | `Macro Settings` dialog. Set macro security to `Notifications for all macros`. Restart Outlook.


## Considerations

Outlook also allows to set macro security to `Enable all macro's`. Even though you will not be bothered for a single confirmation on Outlook startup, it will reduce security as potentially malicious macro's that might be attached to incoming emails could also get active unnoticed. **DO NOT USE THIS SETTING**, and make sure to select `Notifications for all macros` instead.

