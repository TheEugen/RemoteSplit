# 🖥️ RemoteSplit

A Windows C++ desktop application for remote PC control, built on top of the Windows RDP Shared API (RDPSA). The two peers run the same binary — one acts as the sharer (server), the other as the viewer (client). A connection string is exchanged via a shared XML file, which the viewer then reads to establish the remote session.

> **Current state:** The core function is working — a Windows PC can be remote-controlled through a shared file. Further features are planned (see Roadmap).

---

## How It Works

The connection flow is entirely file-based — no direct network address exchange is needed between the two users:

1. **Sharer** clicks "Server" → "Start Sharing Session"
2. The app creates an RDP sharing session via `IRDPSRAPISharingSession`, generates an invitation, and extracts the connection string
3. A save-file dialog opens — the sharer saves the connection string to an `.xml` file and shares it with the viewer (e.g. via cloud storage, USB, messaging)
4. **Viewer** clicks "Client" → "Connect"
5. An open-file dialog prompts the viewer to select the `.xml` file
6. The connection string is read from the file and passed to `IRDPSRAPIViewer::Connect()`
7. The remote desktop stream is rendered in the viewer window via an embedded ActiveX control

---

## Features

- **Win32 GUI** — Native Windows application with a simple button-based UI; switches layout dynamically between Server and Client modes
- **RDP Shared API** — Uses `IRDPSRAPISharingSession` (sharer) and `IRDPSRAPIViewer` (viewer) COM interfaces from the Windows RDPSA
- **File-based connection handshake** — Connection string exchanged via a user-chosen `.xml` file; no manual IP entry required
- **Configurable shared region** — The sharer sets the desktop region to capture via `SetDesktopSharedRect`
- **ActiveX stream rendering** — The viewer embeds the RDP ActiveX control (`{32BE5ED2-5C86-480F-A914-0FF8885A1B3F}`) to display the remote stream in-window
- **Session event handling** — `RDPSessionEvents` implements `_IRDPSessionEvents` to receive and handle session lifecycle events
- **Output log** — Both modes display a scrollable log panel for connection status messages

---

## Project Structure

```
RemoteSplit/
├── RemoteSplit.cpp       # Entry point, Win32 window + message loop, button dispatch
├── Sharer.cpp/.h         # RDP sharing session creation, invitation, XML file write
├── Viewer.cpp/.h         # RDP viewer connection, XML file read, ActiveX stream rendering
├── RDPSessionEvents.h    # COM event sink for _IRDPSessionEvents
├── Utils.h               # GUI helpers, BStr/Variant wrappers, layout switchers
├── associated.cpp/.h     # ActiveX host window registration (AXRegister)
└── Resource.h            # Control IDs and resource definitions
```

---

## Getting Started

### Prerequisites

- Windows (uses RDPSA COM interfaces, Win32 API, and ActiveX)
- Visual Studio (MSVC)
- The Windows RDP Shared API is available on Windows Vista and later

### Build

```bash
git clone https://github.com/TheEugen/RemoteSplit.git
```

Open in Visual Studio and build. No external dependencies — all APIs used (`rdpapi.h`, `atlbase.h`, COM) are part of the Windows SDK.

---

## Roadmap

- [ ] Automatic `.xml` extension enforcement in the save dialog
- [ ] Stop sharing session support
- [ ] Disconnect support on the viewer side
- [ ] Encryption for the connection string / session
- [ ] Support for multiple simultaneous viewers

---

## License

This project is licensed under the [MIT License](LICENSE).