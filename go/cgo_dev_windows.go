//go:build office_oxide_dev && windows

// Dev-mode linker flags for Windows CGo builds.
// Enable with `go build -tags office_oxide_dev ./...`, after running:
//
//	cargo build --release --lib --target x86_64-pc-windows-gnu
//	copy target\x86_64-pc-windows-gnu\release\liboffice_oxide.a target\release\
package officeoxide

// #cgo CFLAGS: -I${SRCDIR}/../include/office_oxide_c
// #cgo LDFLAGS: -L${SRCDIR}/../target/release -loffice_oxide -lws2_32 -lntdll -luserenv
import "C"
