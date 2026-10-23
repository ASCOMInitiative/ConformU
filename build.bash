#!/bin/bash
set -e

# ------------------------------
# VALIDATE THE SUPPLIED BUILD VERSION
# ------------------------------
param="$1"
if [[ "$param" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "Valid version supplied: $param"
else
    echo "Invalid version supplied, script ending: $1"
    exit 1
fi

# ------------------------------
# CONSTANTS
# ------------------------------
HOME="/home/peter" # Root folder on the WSL file system
COMPRESSION_TYPE="gz" # also works with xz

# ------------------------------
# CONFIGURATION
# ------------------------------
REPO_URL="https://github.com/ASCOMInitiative/ConformU.git"

# Linux file system
APPLICATIONNAME="conformu" # Must be lower case!
ROOTFOLDER="$HOME/$APPLICATIONNAME"
CHECKOUTFOLDER="$ROOTFOLDER/checkout"
PUBLISHWORKFOLDER="$ROOTFOLDER/publishwork"
PROJECT="$CHECKOUTFOLDER/ConformU/ConformU.csproj" # Name and location of the project file to be built

# Windows file system
WINDOWSPROJECTFOLDER="/mnt/j/ConformU" # Location where the solution / project code is located
WINDOWSTARFOLDER="$WINDOWSPROJECTFOLDER/publish" # Location on the Windows file system to which the TAR files will be copied

# ------------------------------
# PREPARE WORKDIR
# ------------------------------
echo "Preparing work directories: $CHECKOUTFOLDER and $PUBLISHWORKFOLDER"
rm -rf "$CHECKOUTFOLDER"
mkdir -p "$CHECKOUTFOLDER"
rm -rf "$PUBLISHWORKFOLDER"
mkdir -p "$PUBLISHWORKFOLDER"

cd "$CHECKOUTFOLDER"

echo "Publishing project: $PROJECT"

# ------------------------------
# CLONE REPOSITORY
# ------------------------------
echo "Cloning repository..."
git clone "$REPO_URL" "$CHECKOUTFOLDER"

# ------------------------------
# PUBLISH TARGETS
#     "linux-arm"
#    "linux-arm64"
# ------------------------------
RIDS=(
    "linux-x64"
    "linux-arm"
    "linux-arm64"
)

for RID in "${RIDS[@]}"; do
    OUTDIR="$PUBLISHWORKFOLDER/publish-$RID"
    echo "Publish folder: $OUTDIR"    
    mkdir "$OUTDIR"
    
    TARFILE="$PUBLISHWORKFOLDER/$APPLICATIONNAME.$1.$RID.tar".$COMPRESSION_TYPE
    echo "Tar file: $TARFILE"    
    
    dotnet publish "$PROJECT" \
        -c Release \
        -r "$RID" \
        --self-contained true \
        -p:PublishSingleFile=true \
        -p:PublishReadyToRun=true \
        -p:PublishReadyToRunShowWarnings=true \
        -p:UseAppHost=true \
        -p:TargetFrameworks=net10.0 \
        -p:TargetFramework=net10.0 \
        -o "$OUTDIR"

    cd "$OUTDIR"
    
    echo "Starting tar..."
    time tar --create --verbose --$COMPRESSION_TYPE --file="$TARFILE" *
    echo "Completed tar"

done

cp -f "$PUBLISHWORKFOLDER"/*.$COMPRESSION_TYPE "$WINDOWSTARFOLDER"

echo "All builds complete."
