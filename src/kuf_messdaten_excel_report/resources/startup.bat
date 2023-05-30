start cmd "/K cd C:\script-server && py launcher.py"
start cmd "/K cd C:\repos\another-dauerauswertung && .\.venv\Scripts\activate & cd .\django_dauerauswertung & py ./fun_with_waitress.py"
echo 'start cmd "/K cd C:\repos\another-dauerauswertung && .\.venv\Scripts\activate && cd .\django_dauerauswertung && py manage.py runscript file_watch"'
start cmd "/K cd C:\repos\kufi-django && .\.venv\Scripts\activate & cd .\services\messdaten_processing & py ./watcher.py"
start cmd "/K cd C:\repos\another-dauerauswertung && .\.venv\Scripts\activate && cd .\django_dauerauswertung && py ./services/fun_with_dauerauswertung.py"
start cmd "/K cd C:\repos\viewmes-dash && .venv\Scripts\activate && py main.py"