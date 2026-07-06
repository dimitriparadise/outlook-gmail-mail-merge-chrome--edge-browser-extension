from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "store-assets" / "popup-screenshot-harness.html"


MOCK = """
<script>
  const __mailMergeState = {};
  window.chrome = {
    storage: {
      local: {
        async get(key) {
          if (typeof key === 'string') return { [key]: __mailMergeState[key] };
          return { ...__mailMergeState };
        },
        async set(next) {
          Object.assign(__mailMergeState, next);
        },
        async remove(key) {
          delete __mailMergeState[key];
        }
      }
    },
    tabs: {
      create(_options, callback) {
        callback && callback();
      }
    },
    runtime: {}
  };
</script>
"""


def main():
    html = (ROOT / "popup.html").read_text()
    html = html.replace('<link rel="stylesheet" href="popup.css" />', '<link rel="stylesheet" href="/popup.css" />')
    html = html.replace('<script src="popup.js"></script>', MOCK + '<script src="/popup.js"></script>')
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(html)
    print(OUT)


if __name__ == "__main__":
    main()
