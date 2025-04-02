# Define the root directory (current directory)
$rootDir = "."

# Define the directory names
$directories = @(
    "models",
    "src",
    "tables"
)

# Loop through the directory names and create them
foreach ($dir in $directories) {
    $fullPath = Join-Path -Path $rootDir -ChildPath $dir

    # Check if the directory already exists
    if (-not (Test-Path -Path $fullPath -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $fullPath -Force
            Write-Host "Created directory: $fullPath"
        }
        catch {
            Write-Host "Error creating directory $($fullPath): $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "Directory already exists: $fullPath"
    }
}

Write-Host "Directory structure creation complete."