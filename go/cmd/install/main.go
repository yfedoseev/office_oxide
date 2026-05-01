// Install fetches the office_oxide native library for the current platform
// and prints the CGO_CFLAGS / CGO_LDFLAGS the caller should export.
//
// Usage:
//
//	go run github.com/yfedoseev/office_oxide/go/cmd/install@latest
//	go run github.com/yfedoseev/office_oxide/go/cmd/install@latest -prefix /usr/local
//
// Flags:
//
//	-prefix  Install the header and library under this prefix (default: ~/.office_oxide).
//	-version Release version to fetch (default: the office_oxide version this installer was built against).
//
// After running, export the flags the installer prints, or pass them in a
// go build invocation:
//
//	go build -x ./...  # with CGO_CFLAGS / CGO_LDFLAGS already in env
package main

import (
	"archive/tar"
	"compress/gzip"
	"errors"
	"flag"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"runtime"
	"strings"
)

// Bumped in lockstep with the Rust crate.
const defaultVersion = "0.1.1"

const releaseBase = "https://github.com/yfedoseev/office_oxide/releases/download"

type target struct {
	assetBase string // matches release.yml artifact_name
	libName   string // filename under lib/ in the archive
}

func resolveTarget() (target, error) {
	switch runtime.GOOS + "/" + runtime.GOARCH {
	case "linux/amd64":
		return target{"native-linux-x86_64", "liboffice_oxide.so"}, nil
	case "linux/arm64":
		return target{"native-linux-aarch64", "liboffice_oxide.so"}, nil
	case "darwin/amd64":
		return target{"native-macos-x86_64", "liboffice_oxide.dylib"}, nil
	case "darwin/arm64":
		return target{"native-macos-aarch64", "liboffice_oxide.dylib"}, nil
	case "windows/amd64":
		return target{"native-windows-x86_64", "office_oxide.dll"}, nil
	case "windows/arm64":
		return target{"native-windows-aarch64", "office_oxide.dll"}, nil
	default:
		return target{}, fmt.Errorf(
			"unsupported platform %s/%s — build from source in the office_oxide monorepo",
			runtime.GOOS, runtime.GOARCH,
		)
	}
}

func main() {
	var prefix, version string
	flag.StringVar(&prefix, "prefix", "", "install prefix (default: ~/.office_oxide)")
	flag.StringVar(&version, "version", defaultVersion, "release version")
	flag.Parse()

	if prefix == "" {
		home, err := os.UserHomeDir()
		must(err, "resolve home dir")
		prefix = filepath.Join(home, ".office_oxide")
	}

	tgt, err := resolveTarget()
	must(err, "resolve target")

	ext := ".tar.gz"
	if runtime.GOOS == "windows" {
		ext = ".zip"
	}
	url := fmt.Sprintf("%s/v%s/%s-%s%s", releaseBase, version, tgt.assetBase, version, ext)
	fmt.Fprintf(os.Stderr, "Fetching %s\n", url)

	body, err := httpGet(url)
	must(err, "download asset")
	defer body.Close()

	libDir := filepath.Join(prefix, "lib")
	includeDir := filepath.Join(prefix, "include")
	must(os.MkdirAll(libDir, 0o755), "mkdir lib")
	must(os.MkdirAll(includeDir, 0o755), "mkdir include")

	if runtime.GOOS == "windows" {
		must(errors.New("Windows installer not wired into this minimal helper — "+
			"download the .zip from "+url+" by hand and extract it under "+prefix), "unpack")
	} else {
		must(untarGz(body, prefix), "extract")
	}

	fmt.Printf("\n# office_oxide %s installed under %s\n", version, prefix)
	fmt.Printf("# Export the following before running `go build`:\n\n")
	fmt.Printf("export CGO_CFLAGS=%q\n", "-I"+filepath.Join(prefix, "include", "office_oxide_c"))
	fmt.Printf("export CGO_LDFLAGS=%q\n",
		strings.Join([]string{
			"-L" + libDir,
			"-loffice_oxide",
			"-Wl,-rpath," + libDir,
		}, " "),
	)
}

func httpGet(url string) (io.ReadCloser, error) {
	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	if resp.StatusCode != http.StatusOK {
		resp.Body.Close()
		return nil, fmt.Errorf("HTTP %d for %s", resp.StatusCode, url)
	}
	return resp.Body, nil
}

func untarGz(r io.Reader, dst string) error {
	gz, err := gzip.NewReader(r)
	if err != nil {
		return err
	}
	defer gz.Close()
	tr := tar.NewReader(gz)
	for {
		hdr, err := tr.Next()
		if errors.Is(err, io.EOF) {
			return nil
		}
		if err != nil {
			return err
		}
		target := filepath.Join(dst, filepath.Clean(hdr.Name))
		if !strings.HasPrefix(target, filepath.Clean(dst)+string(os.PathSeparator)) && target != filepath.Clean(dst) {
			return fmt.Errorf("refusing to extract %q outside %q", hdr.Name, dst)
		}
		switch hdr.Typeflag {
		case tar.TypeDir:
			if err := os.MkdirAll(target, 0o755); err != nil {
				return err
			}
		case tar.TypeReg:
			if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
				return err
			}
			f, err := os.OpenFile(target, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.FileMode(hdr.Mode)&0o777|0o600)
			if err != nil {
				return err
			}
			if _, err := io.Copy(f, tr); err != nil {
				f.Close()
				return err
			}
			f.Close()
		}
	}
}

func must(err error, what string) {
	if err != nil {
		fmt.Fprintf(os.Stderr, "install: %s: %v\n", what, err)
		os.Exit(1)
	}
}
