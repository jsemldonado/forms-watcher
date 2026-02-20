# forms-watcher

Stop going to lecture for engagement forms.

Poll Microsoft Forms and get notified when they go live. Works with any Microsoft 365 org.

## Setup

```
uv sync
forms-watcher auth
```

## Usage

```
forms-watcher add https://forms.office.com/r/xxx https://forms.office.com/r/yyy
forms-watcher poll            # poll every 5s (default)
forms-watcher poll --interval 10
forms-watcher status          # check once and exit
forms-watcher list
forms-watcher remove Pivot    # by name, short code, index, or URL
forms-watcher clear
```

Auth tokens refresh automatically and last 90 days.
