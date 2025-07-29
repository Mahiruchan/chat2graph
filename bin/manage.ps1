<#
.SYNOPSIS
A PowerShell script to manage the Chat2Graph application on Windows.

.DESCRIPTION
This script provides functionality to start, stop, restart, check the status, and build the Chat2Graph application.
It is designed to be a replacement for the original .sh scripts for use in a PowerShell environment.

.PARAMETER Command
The action to perform. Must be one of: start, stop, restart, status, build.

.EXAMPLE
# Start the application
.\manage.ps1 -Command start

.EXAMPLE
# Stop the application
.\manage.ps1 -Command stop

.EXAMPLE
# Check the application status
.\manage.ps1 -Command status

.EXAMPLE
# Restart the application
.\manage.ps1 -Command restart

.EXAMPLE
# Build the application (install dependencies and build frontend)
.\manage.ps1 -Command build
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('start', 'stop', 'restart', 'status', 'build')]
    [string]$Command
)

# --- Script Configuration ---

# Set the script's base directory to its location
$script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $script:BaseDir

# Project root is one level above the 'bin' directory
$script:ProjectRoot = (Resolve-Path ..).Path

# Log directory configuration
$script:LogDir = Join-Path -Path $HOME -ChildPath ".chat2graph\logs"
$script:ServerLogFile = Join-Path -Path $script:LogDir -ChildPath "server.log"
$script:McpLogFile = Join-Path -Path $script:LogDir -ChildPath "mcp.log"

# Main application entry point
$script:BootstrapScript = Join-Path -Path $script:ProjectRoot -ChildPath "app\server\bootstrap.py"

# MCP (Multi-Capability Program) tools configuration
# Format: A custom object with Name, Port, and Command properties.
$script:McpToolsConfig = @(
    [pscustomobject]@{
        Name    = "playwright"
        Port    = 8931
        Command = "npx @playwright/mcp@latest --isolated"
    }
    # Add other tools here if needed
)

# Lock file path
$script:LockFile = Join-Path -Path $env:TEMP -ChildPath "chat2graph.lock"

# --- Helper Functions ---

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Get-ServerProcess {
    <#
    .SYNOPSIS
    Finds the running Python server process.
    #>
    $processes = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe'" | Where-Object {
        $_.CommandLine -like "*$($script:BootstrapScript)*"
    }
    return $processes
}

function Start-McpTools {
    Write-Log "Starting MCP servers..." "Cyan"
    
    # Ensure log directory exists
    if (-not (Test-Path -Path $script:LogDir)) {
        New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $newMcpLogPath = Join-Path -Path $script:LogDir -ChildPath "mcp_$($timestamp).log"
    
    foreach ($config in $script:McpToolsConfig) {
        Write-Log "Handling MCP tool: $($config.Name)" "Gray"
        
        # Check if port is in use
        $connection = Get-NetTCPConnection -LocalPort $config.Port -State Listen -ErrorAction SilentlyContinue
        if ($connection) {
            Write-Log "Port $($config.Port) is already in use. Assuming $($config.Name) is running." "Yellow"
            continue
        }
        
        # Start the process
        $commandParts = $config.Command.Split(' ')
        $executable = $commandParts[0]
        $arguments = ($commandParts | Select-Object -Skip 1) -join ' '
        $arguments += " --port $($config.Port)"
        
        try {
            # Construct the full command string for cmd.exe, ensuring proper quoting
            # The entire command string including redirection needs to be quoted for cmd.exe /c
            $fullCommand = "`"$executable`" $arguments"
            $cmdArguments = "/c `"$fullCommand > `"$newMcpLogPath`" 2>&1`""
            
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArguments -WorkingDirectory $script:ProjectRoot -NoNewWindow -PassThru
            Start-Sleep -Seconds 2 # Wait a moment for the process to stabilize
            
            if (Get-Process -Id $process.Id -ErrorAction SilentlyContinue) {
                Write-Log "$($config.Name) MCP tool started successfully! (PID: $($process.Id))" "Green"
            } else {
                throw "Process failed to start."
            }
        }
        catch {
            Write-Log "Failed to start $($config.Name) MCP tool. Check the log for details: $newMcpLogPath" "Red"
        }
    }
    Write-Log "MCP tools logs are being written to $newMcpLogPath" "Gray"
}

function Stop-McpTools {
    # This function is included for completeness, but the original stop script intentionally skips this.
    Write-Log "Stopping MCP servers..." "Cyan"
    foreach ($config in $script:McpToolsConfig) {
        $connection = Get-NetTCPConnection -LocalPort $config.Port -State Listen -ErrorAction SilentlyContinue
        if ($connection) {
            $processId = $connection.OwningProcess
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
            Write-Log "Stopped $($config.Name) MCP tool (PID: $processId)." "Green"
        } else {
            Write-Log "$($config.Name) MCP tool not found or not running." "Yellow"
        }
    }
}

function Acquire-Lock {
    param(
        [string]$LockFilePath
    )
    if (Test-Path -Path $LockFilePath) {
        $lockedPid = Get-Content -Path $LockFilePath
        Write-Log "Error: Build is locked by process $lockedPid. If you are sure no build is running, delete $LockFilePath." "Red"
        exit 1
    }
    try {
        Set-Content -Path $LockFilePath -Value $PID
    } catch {
        Write-Log "Error: Failed to acquire lock file $LockFilePath. $_" "Red"
        exit 1
    }
}

function Release-Lock {
    param(
        [string]$LockFilePath
    )
    if (Test-Path -Path $LockFilePath) {
        $lockedPid = Get-Content -Path $LockFilePath
        if ($lockedPid -eq $PID) {
            Remove-Item -Path $LockFilePath -Force
        } else {
            Write-Log "Warning: Lock file $LockFilePath is held by another process ($lockedPid). Not releasing." "Yellow"
        }
    }
}

function Check-Command {
    param(
        [string]$CommandName,
        [string]$ErrorMessage
    )
    if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
        Write-Log "Error: $CommandName is not found. $ErrorMessage" "Red"
        exit 1
    }
    Write-Log "$CommandName found." "Green"
}

function Handle-DependencyConflicts {
    Write-Log "Resolving aiohttp version conflict..." "Cyan"
    $targetAiohttpVersion = "3.12.13"
    try {
        # Capture output and replace ERROR with WARNING
        $output = (pip install --force-reinstall "aiohttp==$targetAiohttpVersion" 2>&1) -replace "ERROR", "WARNING"
        Write-Log $output "Gray"
    } catch {
        Write-Log "Error resolving aiohttp conflict: $_" "Red"
        exit 1
    }
}

function Build-Python {
    param(
        [string]$ProjectRootPath
    )
    Write-Log "Installing Python packages..." "Cyan"
    Set-Location (Join-Path -Path $ProjectRootPath -ChildPath "app")
    try {
        # Ensure poetry is in PATH or activate venv before running this script
        Invoke-Expression "poetry lock"
        Invoke-Expression "poetry install"
        Handle-DependencyConflicts
        Write-Log "Python packages installed successfully." "Green"
    } catch {
        Write-Log "Error installing Python packages: $_" "Red"
        exit 1
    }
    Set-Location $script:BaseDir # Go back to script directory
}

function Build-Web {
    param(
        [string]$ProjectRootPath
    )
    Write-Log "Building web packages..." "Cyan"
    $webDir = Join-Path -Path $ProjectRootPath -ChildPath "web"
    $serverWebDir = Join-Path -Path $ProjectRootPath -ChildPath "app\server\web"

    Set-Location $webDir
    try {
        Invoke-Expression "npm cache clean --force"
        Invoke-Expression "npm install"
        Invoke-Expression "npm run build"
        
        # Remove existing server web directory and copy new build
        if (Test-Path -Path $serverWebDir -PathType Container) {
            Remove-Item -Path $serverWebDir -Recurse -Force
        }
        Copy-Item -Path (Join-Path -Path $webDir -ChildPath "dist") -Destination $serverWebDir -Recurse -Force
        Write-Log "Web packages built and moved successfully." "Green"
    } catch {
        Write-Log "Error building web packages: $_" "Red"
        exit 1
    }
    Set-Location $script:BaseDir # Go back to script directory
}

# --- Main Logic ---

switch ($Command) {
    "start" {
        Write-Log "Attempting to start Chat2Graph server..." "Cyan"
        
        # Pre-check
        $serverProcess = Get-ServerProcess
        if ($serverProcess) {
            Write-Log "Chat2Graph server already started (PID: $($serverProcess.ProcessId))" "Red"
            exit 1
        }
        
        # Start MCP Servers
        Start-McpTools
        
        # Prepare log path
        if (-not (Test-Path -Path $script:LogDir)) {
            New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $newServerLogPath = Join-Path -Path $script:LogDir -ChildPath "server_$($timestamp).log"
        
        # Start Python Server
        Write-Log "Starting Python server in the background..." "Cyan"
        try {
            # Set PYTHONUTF8 environment variable to force UTF-8 encoding for stdout/stderr
            $env:PYTHONUTF8 = "1"

            # Construct the full command string for cmd.exe, ensuring proper quoting
            # The entire command string including redirection needs to be quoted for cmd.exe /c
            # Note: Assumes 'python' is in the PATH.
            # For virtual environments, activate it in the shell before running this script,
            # or specify the full path to the python.exe in the venv.
            $pythonExecutable = "python"
            # If using a virtual environment, you might need to specify the full path:
            # $pythonExecutable = "C:\Users\Mahiru\Desktop\chat2graph\.venv\Scripts\python.exe"

            $pythonArguments = "-u `"$($script:BootstrapScript)`""
            $fullCommand = "`"$pythonExecutable`" $pythonArguments"
            $cmdArguments = "/c `"$fullCommand > `"$newServerLogPath`" 2>&1`""
            
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArguments -WorkingDirectory $script:ProjectRoot -NoNewWindow -PassThru
            
            Start-Sleep -Seconds 3 # Give it time to start
            
            # Post-check
            $serverProcess = Get-ServerProcess
            if ($serverProcess) {
                Write-Log "Chat2Graph server started successfully! (PID: $($serverProcess.ProcessId))" "Green"
                Write-Log "Server logs are being written to $newServerLogPath" "Gray"
            } else {
                throw "Server process not found after startup."
            }
        }
        catch {
            Write-Log "Chat2Graph server failed to start. Check logs in $newServerLogPath" "Red"
            exit 1
        }
    }
    
    "stop" {
        Write-Log "Attempting to stop Chat2Graph server..." "Cyan"
        
        # The original script intentionally does not stop MCP tools.
        # To stop them, you could call Stop-McpTools here.
        Write-Log "Note: MCP servers are not stopped by this command, by design." "Yellow"

        $serverProcess = Get-ServerProcess
        if ($serverProcess) {
            try {
                Stop-Process -Id $serverProcess.ProcessId -Force
                Write-Log "Chat2Graph server stopped successfully!" "Green"
            } catch {
                Write-Log "Failed to stop Chat2Graph server (PID: $($serverProcess.ProcessId)). It may require manual termination." "Red"
            }
        } else {
            Write-Log "Chat2Graph server not found or already stopped." "Yellow"
        }
    }
    
    "status" {
        $serverProcess = Get-ServerProcess
        if ($serverProcess) {
            Write-Log "Chat2Graph server is RUNNING (PID: $($serverProcess.ProcessId))" "Green"
        } else {
            Write-Log "Chat2Graph server is STOPPED" "Red"
        }
    }
    
    "restart" {
        Write-Log "Restarting Chat2Graph server..." "Cyan"
        # Execute the 'stop' and 'start' logic from this script.
        & $MyInvocation.MyCommand.Definition -Command stop
        Start-Sleep -Seconds 2
        & $MyInvocation.MyCommand.Definition -Command start
    }

    "build" {
        Write-Log "Starting Chat2Graph build process..." "Cyan"
        Acquire-Lock -LockFilePath $script:LockFile
        try {
            Check-Command -CommandName "python" -ErrorMessage "Python is required. Please install it and ensure it's in your PATH."
            Check-Command -CommandName "pip" -ErrorMessage "pip is required. Please install it and ensure it's in your PATH."
            Check-Command -CommandName "poetry" -ErrorMessage "Poetry is required. Run 'pip install poetry' and retry."
            Check-Command -CommandName "node" -ErrorMessage "Node.js is required. Please install it and ensure it's in your PATH."
            Check-Command -CommandName "npm" -ErrorMessage "npm is required. Please install it and ensure it's in your PATH."

            Build-Python -ProjectRootPath $script:ProjectRoot
            Build-Web -ProjectRootPath $script:ProjectRoot

            Write-Log "Build success !" "Green"
        } finally {
            Release-Lock -LockFilePath $script:LockFile
        }
    }
}
