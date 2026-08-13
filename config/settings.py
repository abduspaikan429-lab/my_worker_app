"""Global application settings and configuration."""

# Worker data storage mode: "json" or "sqlite"
WORKER_STORAGE = "json"

# Enable automatic JSON backup for rollback safety
ENABLE_JSON_BACKUP = True

# SQLite database file path
SQLITE_DB_PATH = "data/worker.db"

# JSON state file path (Primary source before migration / Rollback source)
JSON_STATE_FILE = "data/master_state.json"

# JSON backup file path
JSON_BACKUP_FILE = "data/master_state_backup.json"
