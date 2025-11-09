#!/bin/bash
cd DocumentAssemblerSdk.Tests
dotnet test --filter "FullyQualifiedName~DA272" --logger "console;verbosity=detailed" 2>&1 | grep -A 20 "DA272.*Premium.*FAIL"
