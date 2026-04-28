//go:build office_oxide_dev

// Dev-mode linker flags used when building inside the office_oxide monorepo.
// Enable with `go build -tags office_oxide_dev ./...`, after running:
//
//     cargo build --release --lib
//
// Consumers outside the monorepo should unset this tag and export
// CGO_CFLAGS / CGO_LDFLAGS themselves (or use a future installer script).

package officeoxide

// #cgo CFLAGS: -I${SRCDIR}/../include/office_oxide_c
// #cgo LDFLAGS: -L${SRCDIR}/../target/release -loffice_oxide -Wl,-rpath,${SRCDIR}/../target/release -lm -ldl -lpthread
import "C"
