try {
    // Code to parse test suites
    const testSuites = parseTestSuites();
    // Code to run tests
    const testResults = runTests(testSuites);
    // Code to handle test results
    handleTestResults(testResults);
} catch (error) {
    // Handle the error
    console.error('Error parsing test discovery output:', error);
    // Provide a more informative error message
    console.error('Invalid test discovery output. Please check the test suite configuration.');
}Test loading failed: Error: Invalid test discovery output!

Error: Invalid test discovery output!

    at parseTestSuites (/Users/edefrutos/.vscode-insiders/extensions/littlefoxteam.vscode-python-test-adapter-0.8.2/out/src/pytest/pytestTestCollectionParser.js:14:15)
    at PytestTestRunner.<anonymous> (/Users/edefrutos/.vscode-insiders/extensions/littlefoxteam.vscode-python-test-adapter-0.8.2/out/src/pytest/pytestTestRunner.js:60:76)
    at Generator.next (<anonymous>)
    at fulfilled (/Users/edefrutos/.vscode-insiders/extensions/littlefoxteam.vscode-python-test-adapter-0.8.2/node_modules/tslib/tslib.js:114:62)


Stacktrace: 
    at S.syncTopLevel (/Users/edefrutos/.vscode-insiders/extensions/ms-vscode.test-adapter-converter-0.2.1/out/extension.js:6:3)
    at Rh.value (/Users/edefrutos/.vscode-insiders/extensions/ms-vscode.test-adapter-converter-0.2.1/out/extension.js:3:4541)
    at P.B (file:///Applications/Visual%20Studio%20Code%20-%20Insiders.app/Contents/Resources/app/out/vs/workbench/api/node/extensionHostProcess.js:27:2375)
    at P.fire (file:///Applications/Visual%20Studio%20Code%20-%20Insiders.app/Contents/Resources/app/out/vs/workbench/api/node/extensionHostProcess.js:27:2593)
    at Rh.value (/Users/edefrutos/.vscode-insiders/extensions/hbenl.vscode-test-explorer-2.22.1/out/hub/testHub.js:53:43)
    at P.B (file:///Applications/Visual%20Studio%20Code%20-%20Insiders.app/Contents/Resources/app/out/vs/workbench/api/node/extensionHostProcess.js:27:2375)
    at P.C (file:///Applications/Visual%20Studio%20Code%20-%20Insiders.app/Contents/Resources/app/out/vs/workbench/api/node/extensionHostProcess.js:27:2445)
    at P.fire (file:///Applications/Visual%20Studio%20Code%20-%20Insiders.app/Contents/Resources/app/out/vs/workbench/api/node/extensionHostProcess.js:27:2662)
    at PythonTestAdapter.<anonymous> (/Users/edefrutos/.vscode-insiders/extensions/littlefoxteam.vscode-python-test-adapter-0.8.2/out/src/pythonTestAdapter.js:89:35)
    at Generator.throw (<anonymous>)
    at rejected (/Users/edefrutos/.vscode-insiders/extensions/littlefoxteam.vscode-python-test-adapter-0.8.2/node_modules/tslib/tslib.js:115:69)