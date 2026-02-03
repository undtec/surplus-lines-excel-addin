#!/usr/bin/env node
/**
 * Unload script for macOS Excel Add-in
 * Removes manifest from Excel's wef directory
 */

const fs = require('fs');
const path = require('path');
const os = require('os');

const ADDIN_ID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890';

// Get the wef directory based on platform
function getWefDirectory() {
    const platform = os.platform();
    const homeDir = os.homedir();

    if (platform === 'darwin') {
        // macOS - Excel from App Store (sandboxed)
        const sandboxedPath = path.join(
            homeDir,
            'Library/Containers/com.microsoft.Excel/Data/Documents/wef'
        );

        // macOS - Excel from Office installer (non-sandboxed)
        const nonSandboxedPath = path.join(
            homeDir,
            'Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef'
        );

        // Try sandboxed path first (App Store version)
        if (fs.existsSync(path.dirname(sandboxedPath))) {
            return sandboxedPath;
        }

        // Fall back to non-sandboxed path
        return nonSandboxedPath;
    } else if (platform === 'win32') {
        // Windows
        return path.join(
            process.env.LOCALAPPDATA || path.join(homeDir, 'AppData/Local'),
            'Microsoft/Office/16.0/Wef'
        );
    }

    throw new Error(`Unsupported platform: ${platform}`);
}

function unload() {
    const wefDir = getWefDirectory();
    const manifestPath = path.join(wefDir, `${ADDIN_ID}.manifest.xml`);

    console.log('Removing Office Add-in...');
    console.log(`  Path: ${manifestPath}`);

    if (!fs.existsSync(manifestPath)) {
        console.log('  Add-in manifest not found (already removed)');
        return;
    }

    try {
        fs.unlinkSync(manifestPath);
        console.log('✓ Add-in manifest removed successfully!');
        console.log('');
        console.log('Restart Excel to complete the unload.');
    } catch (error) {
        console.error(`Error removing manifest: ${error.message}`);
        process.exit(1);
    }
}

unload();
