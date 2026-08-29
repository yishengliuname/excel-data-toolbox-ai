$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
Push-Location $workspaceRoot
try {
    python -m compileall -q $projectRoot
    python -m pytest $projectRoot\tests
    python -m ruff check $projectRoot\data_contracts.py $projectRoot\lineage.py $projectRoot\ai_evaluation.py $projectRoot\task_engine.py $projectRoot\tool_registry.py $projectRoot\engine_router.py $projectRoot\database_connections.py $projectRoot\conversation.py $projectRoot\session_registry.py
}
finally {
    Pop-Location
}
