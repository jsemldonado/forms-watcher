# forms-watcher

Stop going to lecture for engagement forms.

Poll Microsoft Forms and get notified when they go live. Works with any Microsoft 365 org.

## Setup

```
uv sync
uv run forms-watcher auth
```

## Usage

```
uv run forms-watcher add https://forms.office.com/r/xxx https://forms.office.com/r/yyy
uv run forms-watcher poll            # poll every 5s (default)
uv run forms-watcher poll --interval 10
uv run forms-watcher status          # check once and exit
uv run forms-watcher list
uv run forms-watcher remove Pivot    # by name, short code, index, or URL
uv run forms-watcher clear
```

Auth tokens refresh automatically and last 90 days.
