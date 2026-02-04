// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/log;

# Internal function called from native code to print error logs.
#
# This function is invoked by the Java FailSafeUtils when console logging is enabled.
# It uses the Ballerina log module to output errors to the console.
#
# + message - The error message to log
# + err - Optional error object
# + stackTrace - Optional stack trace
# + keyValues - Optional key-value pairs for structured logging (e.g., offendingRow)
isolated function printError(string message, error? err = (), error:StackFrame[]? stackTrace = (),
        *log:KeyValues keyValues) {
    log:printError(message, err, stackTrace, keyValues = keyValues);
}
