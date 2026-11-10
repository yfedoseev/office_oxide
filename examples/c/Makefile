# Build and run the numbered C examples against the library in target/release.
#   cargo build --release --lib
#   make -C examples/c run
ROOT    := $(abspath ../..)
LIBDIR  := $(ROOT)/target/release
CFLAGS  := -std=c11 -D_DEFAULT_SOURCE -Wall -Wextra -Werror -O1 -I$(ROOT)/include/office_oxide_c
LDFLAGS := -L$(LIBDIR) -Wl,-rpath,$(LIBDIR)
LDLIBS  := -loffice_oxide -lpthread

SRCS := $(sort $(wildcard [0-9][0-9]_*.c))
BINS := $(SRCS:%.c=build/%)

all: $(BINS)

build/%: %.c common.h | build
	$(CC) $(CFLAGS) $< -o $@ $(LDFLAGS) $(LDLIBS)

build:
	mkdir -p build

# Each example prints `OK <id>` as its last line and exits 0.
run: all
	@set -e; for b in $(BINS); do timeout 120 $$b; done

clean:
	rm -rf build

.PHONY: all run clean
