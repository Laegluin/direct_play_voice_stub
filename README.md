An implementation of the direct play voice COM interfaces for Windows 10. I built this to make
the Star Trek Legacy LAN multiplayer work (online using GameRanger). No actual voice communication
is implemented. This may also work with other games that depend on direct play voice.

## Installation

- make sure you have DirectX 9.0c installed and the DirectPlay feature enabled. To enable DirectPlay,
  open PowerShell as Administrator, type `dism /Online /Enable-Feature /FeatureName:DirectPlay /All /NoRestart`
  and hit enter
- copy the `direct_play_voice_stub.dll` somewhere safe, for example the game directory
- open PowerShell as Administrator, type `regsvr32 "path-to-dll"` and hit enter

To uninstall, simply run `regsvr32 /u "path-to-dll"` instead.

## How it works

Direct play voice is not available for Windows 10, even when activating the DirectPlay feature. The
original DLL does not work either. This simply implements all COM interfaces provided by direct play
voice.

The implementation does not perform actual work but only returns dummy values or pretends that
calls were successful. Since voice communication is a non-critical feature, this is enough to make
the game work.

## Building

Building requires Rust 1.48 or higher. Simply run `cargo build` for debug and `cargo build --release`
for release builds.

## Contributing

If you have any ideas or improvements, simply open an issue or submit a PR.

Unless you explicitly state otherwise, any contribution intentionally submitted
for inclusion in the work by you, as defined in the Apache-2.0 license, shall be
dual licensed as above, without any additional terms or conditions.

## License

Licensed under either of

- Apache License, Version 2.0
   ([LICENSE-APACHE](LICENSE-APACHE) or <http://www.apache.org/licenses/LICENSE-2.0>)
- MIT license
   ([LICENSE-MIT](LICENSE-MIT) or <http://opensource.org/licenses/MIT>)

at your option.
