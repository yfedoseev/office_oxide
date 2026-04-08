class OfficeOxide < Formula
  desc "The fastest Office document toolkit — extract text from DOCX, XLSX, PPTX, DOC, XLS, PPT"
  homepage "https://github.com/yfedoseev/office_oxide"
  version "{{VERSION}}"
  license any_of: ["MIT", "Apache-2.0"]

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/yfedoseev/office_oxide/releases/download/v{{VERSION}}/office_oxide-macos-aarch64-{{VERSION}}.tar.gz"
      sha256 "{{SHA256_MACOS_ARM}}"
    else
      url "https://github.com/yfedoseev/office_oxide/releases/download/v{{VERSION}}/office_oxide-macos-x86_64-{{VERSION}}.tar.gz"
      sha256 "{{SHA256_MACOS_X86}}"
    end
  end

  on_linux do
    url "https://github.com/yfedoseev/office_oxide/releases/download/v{{VERSION}}/office_oxide-linux-x86_64-musl-{{VERSION}}.tar.gz"
    sha256 "{{SHA256_LINUX_X86}}"
  end

  def install
    bin.install "office-oxide"
    bin.install "office-oxide-mcp"
  end

  test do
    assert_match version.to_s, shell_output("#{bin}/office-oxide --version")
  end
end
