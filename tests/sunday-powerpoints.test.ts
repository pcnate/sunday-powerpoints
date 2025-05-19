import * as lib from '../src/sunday-powerpoints';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as ws from 'windows-shortcuts';
import { randomUUID } from 'crypto';

jest.mock('windows-shortcuts');
jest.setTimeout(20000);

describe('sundaysInMonth', () => {
  it('returns correct Sundays for May 2024', () => {
    expect(lib.sundaysInMonth(5, 2024)).toEqual([5, 12, 19, 26]);
  });

  it('returns correct Sundays for February 2025', () => {
    expect(lib.sundaysInMonth(2, 2025)).toEqual([2, 9, 16, 23]);
  });

  it('returns correct Sundays for a month starting on Sunday', () => {
    expect(lib.sundaysInMonth(9, 2024)).toEqual([1, 8, 15, 22, 29]);
  });
});

describe('getMonthName', () => {
  it('returns "January" for 1', () => {
    expect(lib.getMonthName(1)).toBe('January');
  });

  it('returns "December" for 12', () => {
    expect(lib.getMonthName(12)).toBe('December');
  });
});

describe('checkIfFileExists', () => {
  let statSpy: jest.SpyInstance;
  beforeEach(() => {
    statSpy = jest.spyOn(fs.promises, 'stat');
  });
  afterEach(() => {
    statSpy.mockRestore();
  });
  it('returns true if file exists', async () => {
    statSpy.mockResolvedValueOnce({} as any);
    await expect(lib.checkIfFileExists('some/path')).resolves.toBe(true);
  });
  it('returns false if file does not exist', async () => {
    statSpy.mockRejectedValueOnce(new Error('not found'));
    await expect(lib.checkIfFileExists('some/path')).resolves.toBe(false);
  });
});

describe('resolveToAbsolutePath', () => {
  it('replaces environment variables with their values', () => {
    process.env.TESTVAR = 'myvalue';
    expect(lib.resolveToAbsolutePath('C:/path/%TESTVAR%/file.txt')).toBe('C:/path/myvalue/file.txt');
  });

  it('replaces multiple environment variables', () => {
    process.env.FOO = 'foo';
    process.env.BAR = 'bar';
    expect(lib.resolveToAbsolutePath('C:/%FOO%/%BAR%/baz')).toBe('C:/foo/bar/baz');
  });

  it('replaces missing environment variables with empty string', () => {
    delete process.env.NOT_SET;
    expect(lib.resolveToAbsolutePath('C:/path/%NOT_SET%/file.txt')).toBe('C:/path//file.txt');
  });
});

// describe('fixExistingShortcuts', () => {
//   it('fixes .lnk shortcut target paths to variable (integration, default behavior)', async () => {
//     const guid = randomUUID();
//     const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), `spp-test-${guid}-`));
//     const shortcutPath = path.join(tempDir, `test-shortcut1-${guid}.lnk`);
//     const targetFile = path.join(tempDir, `target1-${guid}.txt`);
//     fs.writeFileSync(targetFile, 'test content', 'utf8');
//     const originalTarget = targetFile;
//     const expectedTarget = originalTarget.replace(/\\/g, '/');

//     // Create a shortcut with a real file path
//     await new Promise((resolve, reject) => {
//       ws.create(shortcutPath, { target: originalTarget, desc: 'Test Shortcut' }, err => {
//         if (err) reject(err); else resolve(undefined);
//       });
//     });

//     // Run the function under test (no rootPath)
//     await lib.fixExistingShortcuts(tempDir);

//     // Query the shortcut to verify the target was fixed
//     const shortcutInfo = await new Promise<any>((resolve, reject) => {
//       ws.query(shortcutPath, (err, opts) => {
//         if (err) reject(err); else resolve(opts);
//       });
//     });

//     expect(shortcutInfo.target.replace(/\\/g, '/')).toBe(expectedTarget);

//     // Clean up
//     fs.rmSync(tempDir, { recursive: true, force: true });
//   });

//   it('fixes .lnk shortcut target paths to a custom rootPath (integration)', async () => {
//     const guid = randomUUID();
//     const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), `spp-test-${guid}-`));
//     const shortcutPath = path.join(tempDir, `test-shortcut2-${guid}.lnk`);
//     const targetFile = path.join(tempDir, `target2-${guid}.txt`);
//     fs.writeFileSync(targetFile, 'test content', 'utf8');
//     const customRoot = `D:/CustomRoot/${guid}`;
//     const originalTarget = targetFile;
//     const expectedTarget = customRoot + `/target2-${guid}.txt`;

//     // Create a shortcut with a real file path
//     await new Promise((resolve, reject) => {
//       ws.create(shortcutPath, { target: originalTarget, desc: 'Test Shortcut' }, err => {
//         if (err) reject(err); else resolve(undefined);
//       });
//     });

//     // Run the function under test with custom rootPath
//     await lib.fixExistingShortcuts(tempDir, customRoot);

//     // Query the shortcut to verify the target was fixed
//     const shortcutInfo = await new Promise<any>((resolve, reject) => {
//       ws.query(shortcutPath, (err, opts) => {
//         if (err) reject(err); else resolve(opts);
//       });
//     });

//     expect(shortcutInfo.target.replace(/\\/g, '/')).toBe(expectedTarget);

//     // Clean up
//     fs.rmSync(tempDir, { recursive: true, force: true });
//   });
// });

// describe('createShortcut (integration)', () => {
//   it('creates a .lnk shortcut file with correct target and description', async () => {
//     const guid = randomUUID();
//     const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), `spp-create-test-${guid}-`));
//     const shortcutPath = path.join(tempDir, `test-create-shortcut3-${guid}.lnk`);
//     const targetFile = path.join(tempDir, `file3-${guid}.txt`);
//     const description = 'Integration Shortcut Test';
//     // Ensure the target file exists
//     fs.writeFileSync(targetFile, 'test content', 'utf8');
//     // Call the function under test
//     const result = await lib.createShortcut(shortcutPath, description, targetFile);
//     expect(result).toBe(true);
//     // Query the shortcut to verify its properties
//     const shortcutInfo = await new Promise<any>((resolve, reject) => {
//       ws.query(shortcutPath, (err, opts) => {
//         if (err) reject(err); else resolve(opts);
//       });
//     });
//     // The target should match the real path, not the variable
//     expect(shortcutInfo.target.replace(/\\/g, '/')).toBe(targetFile.replace(/\\/g, '/'));
//     expect(shortcutInfo.desc).toBe(description);
//     // Clean up
//     fs.rmSync(tempDir, { recursive: true, force: true });
//   });
// });

// describe('queryOptions (integration)', () => {
//   it('returns the correct options for a real shortcut file', async () => {
//     const guid = randomUUID();
//     const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), `spp-query-test-${guid}-`));
//     const shortcutPath = path.join(tempDir, `test-query-shortcut4-${guid}.lnk`);
//     const targetFile = path.join(tempDir, `file4-${guid}.txt`);
//     const description = 'Query Shortcut Integration Test';
//     // Ensure the target file exists
//     fs.writeFileSync(targetFile, 'test content', 'utf8');
//     // Create the shortcut
//     await new Promise((resolve, reject) => {
//       ws.create(shortcutPath, { target: targetFile, desc: description }, err => {
//         if (err) reject(err); else resolve(undefined);
//       });
//     });
//     // Use the function under test
//     const result = await lib.queryOptions(shortcutPath);
//     expect(result).toBeTruthy();
//     expect((result as any).target.replace(/\\/g, '/')).toBe(targetFile.replace(/\\/g, '/'));
//     expect((result as any).desc).toBe(description);
//     // Clean up
//     fs.rmSync(tempDir, { recursive: true, force: true });
//   });
// });
