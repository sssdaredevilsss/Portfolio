Power tools to automate Remote Access.

- Reading user data from operator (Surname, Name)
- Collecting data from AD, getiing Name,Surname, User`s PC name, email etc.
- Generating dynamic users list if more that one match.
- Adding User and PC to remote access groups (or remove)
- Sshing to VPN concentrator to generate Certs for IKEv2 VPN
- Sending certs and instructions by email
- Generates a personalized VPNer.exe to configure the client station. (Windows based, compilled to exe)
You do not need to download the certificate separately, place it in a folder, and enter the user's login
and the name of the PC for setting shortcuts.
All this data is already sewn into the VPNer, which is generated on the fly.
And such set of shortcuts which was chosen at generation (PC / Terminal / All) will be created.
Generation occurs automatically when creating or resending user certificates.
A universal version of VPNer is sent for Device based certificates
- It is not always convenient to send certificates to the user's mail, as it is usually absent
settings. Therefore, in the menu for sending certificates, the item "Current admin mail" has been added, which will be pulled up
the email address of the user on whose behalf Granter was launched.
There is also an option to specify mail manually.
- The software is packed in a convenient installer.
- Added Updater.exe, for auto-update from the repository.

All exe provided only in prodaction repo.


Also tool can revoke and resend exitsting certs and generate VPNer for it.