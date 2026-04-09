# winvd - Windows 11 virtual desktop bindings for Rust

The implementation abstracts the annoying COM API into simple functions.

https://crates.io/crates/winvd
https://github.com/ciantic/VirtualDesktopAccessor/tree/rust/


### Example

```rust
use winvd::{get_desktop_count, listen_desktop_events, switch_desktop, DesktopEvent};

fn main() {
    // Desktop count
    println!("Desktops: {:?}", get_desktop_count().unwrap());

    // Go to second desktop, index = 1
    switch_desktop(1).unwrap();

    // To listen for changes, use crossbeam, mpsc or winit proxy as a sender
    let (tx, rx) = std::sync::mpsc::channel::<DesktopEvent>();
    let _notifications_thread = listen_desktop_events(tx);

    // Keep the _notifications_thread alive for as long as you wish to listen changes
    std::thread::spawn(|| {
        for item in rx {
            println!("{:?}", item);
        }
    });

    // Wait for keypress
    println!("⛔ Press enter to stop");
    let mut input = String::new();
    std::io::stdin().read_line(&mut input).unwrap();
}
```

`listen_desktop_events` also accepts `crossbeam_channel::Sender<T>` and, with the
optional `winit` feature enabled, `winit::event_loop::EventLoopProxy<T>`.

### Notes

```
cargo clean
cargo doc --all-features
cargo build --release --workspace
```
