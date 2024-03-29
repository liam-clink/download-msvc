#!/usr/bin/env python3

import collections.abc
import io
import os
import sys
import shutil
import json
import shutil
import hashlib
import typing
from urllib.error import URLError
import zipfile
import tempfile
import argparse
import subprocess
import urllib.request
from pathlib import Path

OUTPUT = Path("msvc")  # output folder

# other architectures may work or may not - not really tested
HOST = "x64"  # or x86
TARGET = "x64"  # or x86, arm, arm64

MANIFEST_URL = "https://aka.ms/vs/17/release/channel"


def download(url):
    while True:
        try:
            res = urllib.request.urlopen(url)
            return res.read()
        except URLError:
            continue


def download_with_progress(url, check, name, f):
    # Initialize an in-memory binary stream
    data = io.BytesIO()
    while True:
        try:
            # urlopen() returns a child of Request object with added methods getinfo(), geturl(), info()
            res = urllib.request.urlopen(url)
            total = int(res.headers["Content-Length"])
            size = 0
            while True:
                block = res.read(1 << 20)
                if not block:
                    break
                f.write(block)
                data.write(block)
                size += len(block)
                perc = size * 100 // total
                print(f"\r{name} ... {perc}%", end="")
            break
        except URLError:
            continue
    print()
    data = data.getvalue()
    digest = hashlib.sha256(data).hexdigest()
    if check.lower() != digest:
        exit(f"Hash mismatch for f{pkg}")
    return data


# super crappy msi format parser just to find required .cab files
def get_msi_cabs(msi):
    index = 0
    while True:
        index = msi.find(b".cab", index + 4)
        if index < 0:
            return
        yield msi[index - 32 : index + 4].decode("ascii")


def first(items: collections.abc.Iterable, cond: typing.Callable):
    return next((item for item in items if cond(item)))  # , items[0])


### parse command-line arguments

ap = argparse.ArgumentParser()
ap.add_argument(
    "--show-versions",
    const=True,
    action="store_const",
    help="Show available MSVC and Windows SDK versions",
)
ap.add_argument(
    "--accept-license",
    const=True,
    action="store_const",
    help="Automatically accept license",
)
ap.add_argument("--msvc-version", help="Get specific MSVC version")
ap.add_argument("--sdk-version", help="Get specific Windows SDK version")
ap.add_argument(
    "--get-complete-manifest",
    const=True,
    action="store_const",
    help="Download all packages and combine their manifests, used to determine which packages provides which files",
)
args = ap.parse_args()


### get main manifest
print("Downloading main manifest...")
manifest = json.loads(download(MANIFEST_URL))
### download VS manifest
print("Downloading VS manifest...")
vs = first(
    manifest["channelItems"],
    lambda x: x["id"] == "Microsoft.VisualStudio.Manifests.VisualStudio",
)
# There should only be one payload in the first matching channelItem, this should be the VS manifest
payload = vs["payloads"][0]["url"]
vsmanifest = json.loads(download(payload))
# Keys for vsmanifest: 'manifestVersion', 'engineVersion', 'info', 'signers', 'packages', 'deprecate', 'signature'

## Create dictionary of all packages
packages = {}
for p in vsmanifest["packages"]:
    # This is a bizarre idiom of python. What this does is: if the key p["id"].lower() doesn't exist, add it with an empty list as its value
    # then append p to either the pre-existing list value or the newly created empty list
    # setdefault() doesn't just return the value at a key, it returns a "dictionary view" to that value
    # The view only allows limited operations, it's not like a true reference
    # For example, you cannot increment a view of an integer
    # packages.setdefault(p["id"].lower(), []).append(p)
    # The more understandable (although perhaps less idiomatic) way to do this is
    packages.setdefault(p["id"].lower(), [])
    packages[p["id"].lower()].append(p)
# Write packages.json
fp = open("packages.json", "w")
fp.write(json.dumps(packages, sort_keys=True, indent=4, separators=(",", ": ")))
# Write package_names.txt for easier browsing
fp = open("package_names.txt", "w")
for package_key, package_value in packages.items():
    fp.write(package_key + "\n")


# Get comprehensive manifest (expensive!)
def get_package_manifest(package: str) -> dict:
    p = first(packages[package], lambda p: p.get("language") in (None, "en-US"))
    """    p = first(
        packages[package],
        lambda p: (p.get("language") in (None, "en-US"))
        if ("language" in p.keys())
        else None,
    )"""

    # Currently excluding msi payloads, may change later
    if "payloads" not in p.keys() or p["type"] != "Vsix":
        return
    for payload in p["payloads"]:
        with tempfile.TemporaryFile() as f:
            data = download_with_progress(payload["url"], payload["sha256"], package, f)
            # .vsix files are just .zip files with a different extension
            with zipfile.ZipFile(f) as z:
                if "manifest.json" in z.namelist():
                    out = OUTPUT
                    # out.write_bytes(z.read('manifest.json'))
                    return json.loads(z.read("manifest.json"))


if args.get_complete_manifest:
    print("Getting complete manifest, will take a while...")
    complete_manifest = {}
    for package_key, package_value in packages.items():
        complete_manifest[package_key] = get_package_manifest(package_key)
    fp = open("complete_manifest.json", "w")
    fp.write(
        json.dumps(complete_manifest, sort_keys=True, indent=4, separators=(",", ": "))
    )

### find MSVC & WinSDK versions
msvc_versions = {}
sdk_versions = {}
for package_id in packages.keys():
    if package_id.startswith(
        "Microsoft.VisualStudio.Component.VC.".lower()
    ) and package_id.endswith(".x86.x64".lower()):
        package_version = ".".join(package_id.split(".")[4:-2])
        if package_version[0].isnumeric():
            msvc_versions[package_version] = package_id
    elif package_id.startswith(
        "Microsoft.VisualStudio.Component.Windows10SDK.".lower()
    ):  # or package_id.startswith("Microsoft.VisualStudio.Component.Windows11SDK.".lower()):
        package_version = package_id.split(".")[-1]
        if package_version.isnumeric():
            sdk_versions[package_version] = package_id
## Select MSVC and Windows SDK versions
if args.show_versions:
    print("MSVC versions:", " ".join(sorted(msvc_versions.keys())))
    print("Windows SDK versions:", " ".join(sorted(sdk_versions.keys())))
    exit(0)

msvc_version = args.msvc_version or max(sorted(msvc_versions.keys()))
sdk_version = args.sdk_version or max(sorted(sdk_versions.keys()))

if msvc_version in msvc_versions:
    msvc_pid = msvc_versions[msvc_version]
    if msvc_version != ".".join(msvc_pid.split(".")[4:-2]):
        # This shouldn't be possible
        raise ValueError("Input msvc_version is invalid")
else:
    exit(f"Unknown MSVC version: v{args.msvc_version}")

if sdk_version in sdk_versions:
    sdk_package_id = sdk_versions[sdk_version]
else:
    exit(f"Unknown Windows SDK version: v{args.sdk_version}")

print(f"Downloading MSVC v{msvc_version} and Windows SDK v{sdk_version}")


### agree to license

tools = first(
    manifest["channelItems"],
    lambda x: x["id"] == "Microsoft.VisualStudio.Product.BuildTools",
)
resource = first(tools["localizedResources"], lambda x: x["language"] == "en-us")
license = resource["license"]

if not args.accept_license:
    accept = input(f"Do you accept Visual Studio license at {license} ? [y/n] ")
    if accept and accept[0].lower() != "y":
        exit(0)
print("test 1")

## Clear output directory
if OUTPUT.exists():
    delete = input("Output directory already exists, delete? [y/n] ")
    if delete and delete[0].lower != "y":
        sys.exit("Program terminated. Remove output directory to safely proceed.")
    shutil.rmtree(OUTPUT, ignore_errors=False)
OUTPUT.mkdir()

### download MSVC
# select from "microsoft.visualstudio.workload.nativedesktop"
msvc_packages = [
    # MSVC binaries
    f"microsoft.vc.{msvc_version}.tools.host{HOST}.target{TARGET}.base",
    f"microsoft.vc.{msvc_version}.tools.host{HOST}.target{TARGET}.res.base",
    # MSVC headers
    f"microsoft.vc.{msvc_version}.crt.headers.base",
    # MSVC libs
    f"microsoft.vc.{msvc_version}.crt.{TARGET}.desktop.base",
    f"microsoft.vc.{msvc_version}.crt.{TARGET}.store.base",
    # MSVC runtime source
    f"microsoft.vc.{msvc_version}.crt.source.base",
    # ASAN
    f"microsoft.vc.{msvc_version}.asan.headers.base",
    f"microsoft.vc.{msvc_version}.asan.{TARGET}.base",
    # MSVC redist
    # f"microsoft.vc.{msvc_ver}.crt.redist.x64.base",
    # Need vcvars to set up environment variables and such
    f"microsoft.visualstudio.vc.vcvars",
    # These tools include the Auxiliary/Build/Microsoft.VCToolsVersion.default.txt
    # and Auxiliary/Build/Microsoft.VCRedistVersion.default.txt files
    # asked for by the vcvars.bat
    f"microsoft.visualcpp.tools.core.x86",
    # Needed for vsdevcmd.bat etc
    f"microsoft.visualstudio.vc.devcmd",
    ## Optional stuff
    # CPPWinRT, Dev17 is the codename for Visual Studio 22
    f"microsoft.windows.cppwinrt.dev17",
    # Microsoft Foundational Classes
    f"microsoft.visualstudio.component.vc.{msvc_version}.mfc",
]

print("Downloading selected MSVC Packages")
total_download = 0
already_gotten_packages = []


def get_package(package: str):
    global total_download
    global already_gotten_packages

    if package in already_gotten_packages:
        return

    p = first(packages[package], lambda p: p.get("language") in (None, "en-US"))
    if "dependencies" in p.keys():
        dependencies = p["dependencies"]
        for dep_key, dep_val in dependencies.items():
            if type(dep_val) == str:
                get_package(dep_key.lower())
            elif "id" in dep_val.keys():
                get_package(dep_val["id"].lower())

    # Currently excluding msi payloads, may change later
    if "payloads" not in p.keys() or p["type"] != "Vsix":
        return
    for payload in p["payloads"]:
        with tempfile.TemporaryFile() as f:
            data = download_with_progress(payload["url"], payload["sha256"], package, f)
            total_download += len(data)
            # .vsix files are just .zip files with a different extension
            with zipfile.ZipFile(f) as z:
                for name in z.namelist():
                    if name.startswith("Contents/"):
                        out = OUTPUT / Path(name).relative_to("Contents")
                        out.parent.mkdir(parents=True, exist_ok=True)
                        out.write_bytes(z.read(name))
    already_gotten_packages.append(p["id"].lower())


for pkg in msvc_packages:
    get_package(pkg)

### download Windows SDK
sdk_packages = [
    # Windows SDK tools (like rc.exe & mt.exe)
    f"Windows SDK for Windows Store Apps Tools-x86_en-us.msi",
    # Windows SDK headers
    f"Windows SDK for Windows Store Apps Headers-x86_en-us.msi",
    f"Windows SDK Desktop Headers {TARGET}-x86_en-us.msi",
    # Windows SDK libs
    f"Windows SDK for Windows Store Apps Libs-x86_en-us.msi",
    f"Windows SDK Desktop Libs {TARGET}-x86_en-us.msi",
    # Windows SDK Desktop Tools
    f"Windows SDK Desktop Tools x64-x86_en-us.msi",
    # CRT headers & libs
    f"Universal CRT Headers Libraries and Sources-x86_en-us.msi",
    # CRT redist
    # "Universal CRT Redistributable-x86_en-us.msi",
]

with tempfile.TemporaryDirectory() as d:
    dst = Path(d)

    sdk_pkg = packages[sdk_package_id][0]
    sdk_pkg = packages[first(sdk_pkg["dependencies"], lambda x: True).lower()][0]

    payloads = [ele["fileName"] for ele in sdk_pkg["payloads"]]

    fp = open("sdk_payloads.txt", "w")
    for payload in payloads:
        fp.write(payload + "\n")
    fp.close()

    msi = []
    cabs = []

    # download msi files
    for pkg in sdk_packages:
        payload = first(
            sdk_pkg["payloads"], lambda p: p["fileName"] == f"Installers\\{pkg}"
        )
        msi.append(dst / pkg)
        with open(dst / pkg, "wb") as f:
            data = download_with_progress(payload["url"], payload["sha256"], pkg, f)
            total_download += len(data)
            cabs += list(get_msi_cabs(data))

    # download .cab files
    for pkg in cabs:
        payload = first(
            sdk_pkg["payloads"], lambda p: p["fileName"] == f"Installers\\{pkg}"
        )
        with open(dst / pkg, "wb") as f:
            download_with_progress(payload["url"], payload["sha256"], pkg, f)

    print("Unpacking msi files...")

    # run msi installers
    for m in msi:
        subprocess.check_call(
            ["msiexec.exe", "/a", m, "/quiet", "/qn", f"TARGETDIR={OUTPUT.resolve()}"]
        )


### versions

msvcv = list((OUTPUT / "VC/Tools/MSVC").glob("*"))[0].name
sdkv = list((OUTPUT / "Windows Kits/10/bin").glob("*"))[0].name


# place debug CRT runtime into MSVC folder (not what real Visual Studio installer does... but is reasonable)

dst = str(OUTPUT / "VC/Tools/MSVC" / msvcv / f"bin/Host{HOST}/{TARGET}")

pkg = "microsoft.visualcpp.runtimedebug.14"
dbg = packages[pkg][0]
payload = first(dbg["payloads"], lambda p: p["fileName"] == "cab1.cab")
try:
    with tempfile.TemporaryFile(suffix=".cab", delete=False) as f:
        data = download_with_progress(payload["url"], payload["sha256"], pkg, f)
        total_download += len(data)
    subprocess.check_call(
        ["expand.exe", f.name, "-F:*", dst],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
finally:
    os.unlink(f.name)


### cleanup

# shutil.rmtree(OUTPUT / "Common7", ignore_errors=True)
for f in ["Auxiliary", f"lib/{TARGET}/store", f"lib/{TARGET}/uwp"]:
    shutil.rmtree(OUTPUT / "VC/Tools/MSVC" / msvcv / f)
for f in OUTPUT.glob("*.msi"):
    f.unlink()
for f in ["Catalogs", "DesignTime", f"bin/{sdkv}/chpe", f"Lib/{sdkv}/ucrt_enclave"]:
    shutil.rmtree(OUTPUT / "Windows Kits/10" / f, ignore_errors=True)
for arch in ["x86", "x64", "arm", "arm64"]:
    if arch != TARGET:
        shutil.rmtree(
            OUTPUT / "VC/Tools/MSVC" / msvcv / f"bin/Host{arch}", ignore_errors=True
        )
        shutil.rmtree(OUTPUT / "Windows Kits/10/bin" / sdkv / arch)
        shutil.rmtree(OUTPUT / "Windows Kits/10/Lib" / sdkv / "ucrt" / arch)
        shutil.rmtree(OUTPUT / "Windows Kits/10/Lib" / sdkv / "um" / arch)


### setup.bat
###TODO: change path of vswhere.exe in Common7/Tools/Launch-VsDevShell.ps1 and VsDevCmd.bat
SETUP = f"""@echo off

set ROOT=%~dp0

set MSVC_VERSION={msvcv}
set MSVC_HOST=Host{HOST}
set MSVC_ARCH={TARGET}
set SDK_VERSION={sdkv}
set SDK_ARCH={TARGET}

set MSVC_ROOT=%ROOT%VC\\Tools\\MSVC\\%MSVC_VERSION%
set SDK_INCLUDE=%ROOT%Windows Kits\\10\\Include\\%SDK_VERSION%
set SDK_LIBS=%ROOT%Windows Kits\\10\\Lib\\%SDK_VERSION%

set VCToolsInstallDir=%MSVC_ROOT%\\
set PATH=%MSVC_ROOT%\\bin\\%MSVC_HOST%\\%MSVC_ARCH%;%ROOT%Windows Kits\\10\\bin\\%SDK_VERSION%\\%SDK_ARCH%;%ROOT%Windows Kits\\10\\bin\\%SDK_VERSION%\\%SDK_ARCH%\\ucrt;%PATH%
set INCLUDE=%MSVC_ROOT%\\include;%SDK_INCLUDE%\\ucrt;%SDK_INCLUDE%\\shared;%SDK_INCLUDE%\\um;%SDK_INCLUDE%\\winrt;%SDK_INCLUDE%\\cppwinrt
set LIB=%MSVC_ROOT%\\lib\\%MSVC_ARCH%;%SDK_LIBS%\\ucrt\\%SDK_ARCH%;%SDK_LIBS%\\um\\%SDK_ARCH%
"""

with open(OUTPUT / "setup.bat", "w") as f:
    print(SETUP, file=f)

print(f"Total downloaded: {total_download>>20} MB")
print("Done!")
