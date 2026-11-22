## Objective

Add Conform Universal support for selecting and conformance-testing Alpaca `Aligner` devices through the ASCOM `IAlignerV1` interface and the existing `AlpacaAligner` client. Add a new tester that participates in the existing common lifecycle, but leave each Aligner-specific property and method test as an ordered TODO/remark placeholder. Test `IAlignerV1.Configuration` before all other Aligner-specific properties. The `Configuration` placeholder must describe validation of the returned concrete object against `IAlignerConfigurationV1`, including access to every configuration member.

Support only Alpaca Aligner devices. Do not add a COM facade, COM DriverAccess path, or settings for COM Aligner drivers.

## Relevant Existing Architecture

- `ConformU\ConformanceTestManager.cs` selects one `DeviceTesterBaseClass` subclass in `AssignTestDevice()`. The normal lifecycle is initialise, create, pre-connect checks, connect, common tests, can-properties, pre-run checks, properties, methods, performance, post-run checks, disconnect, and configuration checks.
- Device-specific testers under `ConformU\Conform` derive from `DeviceTesterBaseClass`. They construct an `I*Vn` client in `CreateDevice()`, pass it to `SetDevice(...)`, wait for stabilization, validate the interface version, and dispose it in an override of `Dispose(bool)`.
- The referenced `ASCOM.Alpaca.Components` 4.2.0-rc.1 and `ASCOM.Common.Components` 4.2.0-rc.1 already provide:
  - `ASCOM.Common.DeviceTypes.Aligner`;
  - `ASCOM.Common.DeviceInterfaces.IAlignerV1`;
  - `ASCOM.Common.DeviceInterfaces.IAlignerConfigurationV1`;
  - `ASCOM.Alpaca.Clients.AlpacaAligner`;
  - `ASCOM.Common.DeviceStateClasses.AlignerState`.
- `IAlignerV1` has the properties `Configuration`, `IsMoving`, `XPosition`, `YPosition`, and `ZPosition`, plus methods `FindHome()`, `Halt()`, `Move(double X, double Y, double Z)`, `MoveAbsolute(double X, double Y, double Z)`, and `Sync(double X, double Y, double Z)`.
- `IAlignerConfigurationV1` has 21 read-only properties: `AlignerType`, `BacklashCompensation`, `CanFindHome`, `CanMoveX`, `CanMoveY`, `CanMoveZ`, `CanMoveAbsolute`, `CanSync`, `Units`, `XMaxMove`, `XMaxPosition`, `XMinPosition`, `XResolution`, `YMaxMove`, `YMaxPosition`, `YMinPosition`, `YResolution`, `ZMaxMove`, `ZMaxPosition`, `ZMinPosition`, and `ZResolution`.
- `DeviceTesterBaseClass.CheckCommonMethods()` handles the common IAscomDevice members and `DeviceState`. Its `DeviceState` switch must know the expected operational property names for every device type.
- `ConformU\Pages\DeviceSelect.razor` uses one explicit device-type list for the Alpaca and COM tabs. `ConformU\AlpacaProtocolTestManager.cs` separately dispatches device-specific raw Alpaca protocol tests.
- `Program.GetDeviceTechnology()` already builds its Alpaca URI and COM ProgID patterns from `Devices.DeviceTypeNames()`. Because that Library list now includes Aligner, no parser change is needed.
- `ConformU.Tests` is an existing xUnit v3 test project. Its tests run through the Microsoft.Testing.Platform in-process runner. The current `.NET 10` `dotnet test` VSTest route fails before running tests; use `dotnet run --project ConformU.Tests\ConformU.Tests.csproj ...` with xUnit runner arguments.
- The referenced Library's capability metadata is only partially updated for Aligner. `DeviceCapabilities.LatestInterface` contains Aligner version 1 and `IsValidAscomInterface(Aligner, 1)` is true, but `LatestPlatform6Interface` and `InitialPlatform7Interface` do not contain Aligner. Current calls to `DeviceCapabilities.HasConnectAndDeviceState(Aligner, 1)` and `DeviceCapabilities.IsPlatform7OrLater(Aligner, 1)` throw `KeyNotFoundException`. The application must compensate for this package inconsistency without changing behaviour for other device types.

## Design Decisions

1. Use a single `AlignerTester` in `ConformU\Conform\AlignerTester.cs`, based on the lifecycle pattern in `ObservingConditionsTester`, `FocuserTester`, and `SafetyMonitorTester`.
2. Keep all Aligner-specific conformance checks as executable no-op overrides containing TODO comments. Do not invoke any `IAlignerV1`-specific property or method during this implementation phase.
3. Instantiate `AlpacaAligner` directly. Never reference `ASCOM.Com.DriverAccess`, create an Aligner facade, or add a Windows conditional COM branch.
4. Treat Aligner V1 as a Platform 7 interface for ConformU lifecycle decisions: it has `IAscomDeviceV2`, `Connect`, `Disconnect`, `Connecting`, and `DeviceState`. Centralize this one-device exception in a small capability wrapper so the existing external `DeviceCapabilities` calls are not scattered with Aligner conditionals.
5. Preserve `Configuration` ordering in both the conformance and protocol placeholder lists. The first Aligner-specific placeholder in `CheckProperties()` must be `Configuration`; the remaining IAligner properties follow it.
6. Do not invent a ConformU-owned configuration DTO or adapter. The Library contract says `AlpacaAligner.Configuration` returns `IAlignerConfigurationV1`. A later implemented test must validate the actual runtime object and report a Library/client failure if the contract is not met.
7. Add Aligner to the normal device selector, but disable and avoid selecting the COM tab while Aligner is selected.
8. Route Aligner through both the conformance manager and the Alpaca protocol manager. The protocol-specific tests are placeholders in this phase, just like the conformance-specific tests.
9. Do not change persisted settings. No Aligner-specific safety timeout or feature flag is needed for placeholder-only support.

## Implementation Steps

### Step 1: Add a narrow capability compatibility layer for Aligner

Location:
- Add `ConformU\DeviceCapabilitySupport.cs`.

Symbols:
- `internal static class DeviceCapabilitySupport`
- `HasConnectAndDeviceState(DeviceTypes? deviceType, int interfaceVersion)`
- `IsPlatform7OrLater(DeviceTypes? deviceType, int interfaceVersion)`
- `TryGetLatestPlatform6Interface(DeviceTypes deviceType, out short interfaceVersion)`

Changes:
1. Add the required `using ASCOM.Common;` and `using ASCOM.Common.DeviceInterfaces;` imports.
2. Implement `HasConnectAndDeviceState(...)` so that:
   - for `DeviceTypes.Aligner`, it returns `true` only when `interfaceVersion == 1`;
   - for every other type or null, it delegates exactly to `DeviceCapabilities.HasConnectAndDeviceState(deviceType, interfaceVersion)`.
3. Implement `IsPlatform7OrLater(...)` so that:
   - for `DeviceTypes.Aligner`, it returns `true` only when `interfaceVersion == 1`;
   - for every other type or null, it delegates exactly to `DeviceCapabilities.IsPlatform7OrLater(deviceType, interfaceVersion)`.
4. Implement `TryGetLatestPlatform6Interface(...)` as a guarded lookup into `DeviceCapabilities.LatestPlatform6Interface`. Return `true` and the value when the key exists; otherwise set `interfaceVersion = 0` and return `false`.

Important details:
- Do not catch broad exceptions in these methods. The only Aligner exception is the missing dictionary key, which the explicit type branch avoids. Non-Aligner calls must retain their current failure behavior.
- The wrapper methods intentionally model only the valid Aligner V1 interface. They must return `false` for Aligner interface versions 0, 2, and greater.
- These methods centralize a Library metadata gap. Do not copy Aligner dictionaries into `DeviceCapabilitySupport`.

Must preserve:
- Existing capability results for all non-Aligner device types.
- Existing exception behavior for invalid non-Aligner inputs.

Verification:
- The class compiles on both `net10.0` and `net10.0-windows`.
- Step 7's unit tests prove the Aligner special case and non-Aligner delegation.

### Step 2: Route shared lifecycle capability checks through the wrapper and add Aligner DeviceState members

Location:
- `ConformU\Conform\DeviceTesterBaseClass\DeviceTesterBaseClass.cs`
- `ConformU\AlpacaProtocolTestManager.cs`
- `ConformU\ConformanceTestManager.cs`

Symbols:
- `DeviceTesterBaseClass.CheckCommonMethods()`
- `DeviceTesterBaseClass.Connect()`
- `DeviceTesterBaseClass.Disconnect()`
- capability checks in `DeviceTesterBaseClass`
- `AlpacaProtocolTestManager.TestCommon()`
- `AlpacaProtocolTestManager.TestConnect(IAscomDeviceV2)`
- exception reporting in `ConformanceTestManager.TestDevice(int)`

Changes:
1. In `DeviceTesterBaseClass`, replace every direct call to `DeviceCapabilities.HasConnectAndDeviceState(...)` with `DeviceCapabilitySupport.HasConnectAndDeviceState(...)`.
2. In `DeviceTesterBaseClass.Connect()`, replace the assignment from `DeviceCapabilities.IsPlatform7OrLater(...)` with `DeviceCapabilitySupport.IsPlatform7OrLater(...)`.
3. In the `DeviceState` expected-property switch in `DeviceTesterBaseClass.CheckCommonMethods()`, add:
   ```csharp
   case DeviceTypes.Aligner:
       casedPropertyNames.Add(nameof(IAlignerV1.IsMoving), false);
       casedPropertyNames.Add(nameof(IAlignerV1.XPosition), false);
       casedPropertyNames.Add(nameof(IAlignerV1.YPosition), false);
       casedPropertyNames.Add(nameof(IAlignerV1.ZPosition), false);
       break;
   ```
   Do not add `Configuration` as an expected `DeviceState` member. `AlignerState` exposes the four operational members shown above.
4. In `AlpacaProtocolTestManager.TestCommon()` and `TestConnect(...)`, replace direct `DeviceCapabilities.HasConnectAndDeviceState(...)` calls with the corresponding `DeviceCapabilitySupport` method.
5. In the two `ConformanceTestManager.TestDevice(int)` error-reporting blocks that directly index `DeviceCapabilities.LatestPlatform6Interface[deviceType]`, first call `DeviceCapabilitySupport.TryGetLatestPlatform6Interface(...)`.
   - If it returns `true`, preserve the existing Platform 6 message with the returned value.
   - If it returns `false`, do not index the dictionary and do not emit a Platform 6 latest-version value.
   - Keep the existing `DeviceCapabilities.LatestInterface[deviceType]` message; the Aligner key exists in that dictionary.

Important details:
- Replace all occurrences, not only the calls exercised first by an Aligner run. A missed call in connect, disconnect, common methods, or error reporting can still reintroduce the Aligner `KeyNotFoundException`.
- `DeviceCapabilitySupport` must only change capability selection. Do not replace calls to `DeviceCapabilities.InterfaceHasMember(...)`, `IsValidAscomInterface(...)`, or `LatestInterface`; those already handle Aligner V1 correctly for the use cases here.
- The expected `DeviceState` order is not semantically important, but use the `IsMoving`, `XPosition`, `YPosition`, `ZPosition` order from `AlignerState`.

Must preserve:
- Existing DeviceState validation, casing checks, missing-member reporting, and `TimeStamp` handling.
- Existing connect/disconnect algorithms for every non-Aligner device.
- Existing exception logs and user messages for types that have a Platform 6 entry.

Verification:
- Search the three modified files for direct `DeviceCapabilities.HasConnectAndDeviceState` and `DeviceCapabilities.IsPlatform7OrLater` calls; none should remain.
- Search `ConformanceTestManager.cs` for direct `LatestPlatform6Interface[...]` indexing; none should remain.
- Run an Aligner V1 conformance attempt far enough to execute common capability checks without `KeyNotFoundException`.

### Step 3: Add an Alpaca-only Aligner tester with ordered TODO placeholders

Location:
- Add `ConformU\Conform\AlignerTester.cs`.

Symbols:
- `internal class AlignerTester : DeviceTesterBaseClass`
- `AlignerTester(...)`
- `Dispose(bool)`
- `InitialiseTest()`
- `CreateDevice()`
- `CheckProperties()`
- `CheckMethods()`
- `CheckPerformance()`
- `CheckConfiguration()`

Changes:
1. Import only the packages needed for Alpaca/common ASCOM use, reflection, threading, and base exceptions. Use `ASCOM.Alpaca.Clients`, `ASCOM.Common`, and `ASCOM.Common.DeviceInterfaces`. Do not import `ASCOM.Com.DriverAccess`.
2. Store an `IAlignerV1 aligner`, a `Settings`, and a `ConformLogger`, matching existing tester patterns. Do not store a COM-specific object.
3. Define the constructor with the same dependency signature as other testers:
   ```csharp
   public AlignerTester(ConformConfiguration conformConfiguration,
                        ConformLogger logger,
                        CancellationToken conformCancellationToken,
                        CancellationTokenSource conformCancellationTokenSource,
                        ConformResults conformResults)
   ```
   Call the base constructor with:
   - `hasCanProperties = false`;
   - `hasProperties = true`;
   - `hasMethods = true`;
   - `hasPreRunCheck = false`;
   - `hasPreConnectCheck = false`;
   - `hasPerformanceCheck = true`;
   - `hasPostRunCheck = false`.
4. In `Dispose(bool)`, dispose `aligner`, set it to null, call `base.Dispose(disposing)`, and retain the repository's idempotent `disposedValue` pattern.
5. Override `InitialiseTest()`. Use the same default ASCOM exception numbers used by simple testers, then call `base.InitialiseTest()`. Do not inspect or branch on `settings.ComDevice.ProgId`; Aligner has no COM support.
   Set `ExNotImplemented = 0x80040400`, `ExInvalidValue1` through `ExInvalidValue6 = 0x80040405`, and `ExNotSet1 = 0x80040403` inside the existing `unchecked` pattern.
6. Override `CreateDevice()`:
   - log the configured Alpaca address, port, and device number;
   - accept only `settings.DeviceTechnology == DeviceTechnology.Alpaca`;
   - construct `AlpacaAligner` with the same 15 arguments used by the existing Alpaca constructors: service type, IP address, port, device number, establish timeout, standard timeout, long timeout, `Globals.CLIENT_NUMBER_DEFAULT`, username, password, strict casing, trace logger-or-null, `Globals.USER_AGENT_PRODUCT_NAME`, the executing assembly version as `ToString(4)`, and trust-user-generated-SSL-certificate setting;
   - assign with `SetDevice(aligner, DeviceTypes.Aligner)`;
   - call `SetFullStatus(...)`, `WaitFor(1000, 100)`, and `ValidateInterfaceVersion()`.
   - for COM or any other technology, throw a clear `InvalidValueException` stating that Aligner supports only Alpaca devices.
   - log and rethrow other exceptions in the same style as the existing testers.
7. Override `CheckProperties()` with comments only and no executable Aligner-specific member access. Put these placeholders in this exact order:
   1. `// TODO: Test IAlignerV1.Configuration before all other IAlignerV1 properties.`
   2. `// TODO: Verify Configuration is non-null and that its runtime type is assignable to IAlignerConfigurationV1.`
   3. `// TODO: Verify Configuration.AlignerType is accessible and cannot throw PropertyNotImplementedException.`
   4. `// TODO: Verify Configuration.BacklashCompensation is accessible and cannot throw PropertyNotImplementedException.`
   5. `// TODO: Verify Configuration.CanFindHome is accessible and cannot throw PropertyNotImplementedException.`
   6. `// TODO: Verify Configuration.CanMoveX is accessible and cannot throw PropertyNotImplementedException.`
   7. `// TODO: Verify Configuration.CanMoveY is accessible and cannot throw PropertyNotImplementedException.`
   8. `// TODO: Verify Configuration.CanMoveZ is accessible and cannot throw PropertyNotImplementedException.`
   9. `// TODO: Verify Configuration.CanMoveAbsolute is accessible and cannot throw PropertyNotImplementedException.`
   10. `// TODO: Verify Configuration.CanSync is accessible and cannot throw PropertyNotImplementedException.`
   11. `// TODO: Verify Configuration.Units is accessible and cannot throw PropertyNotImplementedException.`
   12. `// TODO: Test Configuration.XMaxMove for an implemented value or PropertyNotImplementedException.`
   13. `// TODO: Test Configuration.XMaxPosition for an implemented value or PropertyNotImplementedException.`
   14. `// TODO: Test Configuration.XMinPosition for an implemented value or PropertyNotImplementedException.`
   15. `// TODO: Test Configuration.XResolution for an implemented value or PropertyNotImplementedException.`
   16. `// TODO: Test Configuration.YMaxMove for an implemented value or PropertyNotImplementedException.`
   17. `// TODO: Test Configuration.YMaxPosition for an implemented value or PropertyNotImplementedException.`
   18. `// TODO: Test Configuration.YMinPosition for an implemented value or PropertyNotImplementedException.`
   19. `// TODO: Test Configuration.YResolution for an implemented value or PropertyNotImplementedException.`
   20. `// TODO: Test Configuration.ZMaxMove for an implemented value or PropertyNotImplementedException.`
   21. `// TODO: Test Configuration.ZMaxPosition for an implemented value or PropertyNotImplementedException.`
   22. `// TODO: Test Configuration.ZMinPosition for an implemented value or PropertyNotImplementedException.`
   23. `// TODO: Test Configuration.ZResolution for an implemented value or PropertyNotImplementedException.`
   24. `// TODO: Verify the XMaxPosition/XMinPosition, YMaxPosition/YMinPosition, and ZMaxPosition/ZMinPosition implementation states agree within each axis pair.`
   25. `// TODO: Test IAlignerV1.IsMoving.`
   26. `// TODO: Test IAlignerV1.XPosition.`
   27. `// TODO: Test IAlignerV1.YPosition.`
   28. `// TODO: Test IAlignerV1.ZPosition.`
   The comments are placeholders; do not add temporary calls merely to check access. The future concrete test must evaluate `Configuration.GetType()` and verify `typeof(IAlignerConfigurationV1).IsAssignableFrom(...)`; it must not only trust the property's compile-time return type.
8. Override `CheckMethods()` with comments only, in this order:
   - `// TODO: Test IAlignerV1.FindHome().`
   - `// TODO: Test IAlignerV1.Halt().`
   - `// TODO: Test IAlignerV1.Move(double X, double Y, double Z).`
   - `// TODO: Test IAlignerV1.MoveAbsolute(double X, double Y, double Z).`
   - `// TODO: Test IAlignerV1.Sync(double X, double Y, double Z).`
   Do not call any of these methods in this implementation.
9. Override `CheckPerformance()` and make its only executable device-independent action `PerformanceTestCommon(ApplicationCancellationToken)`. Clear the test/action/status using the repository's existing pattern. Do not add Aligner property or method performance loops.
10. Override `CheckConfiguration()` using the common pattern from `FocuserTester`:
   - alert if `settings.TestProperties` is false;
   - alert if `settings.TestMethods` is false;
   - alert if `AllowConnectedTrueAfterDisconnect` is true;
   - log errors through `LogError`/`LogDebug`.

Important details:
- `IAlignerV1.Configuration` is declared mandatory and must be first in the future test ordering. The `IsMoving`, `XPosition`, `YPosition`, and `ZPosition` placeholders must appear after every configuration placeholder.
- The property-level TODOs must distinguish the nine mandatory configuration values (`AlignerType`, `BacklashCompensation`, the six `Can*` values, and `Units`) from the twelve optional numeric values.
- Optional numeric values represent absence by `PropertyNotImplementedException` at the `IAlignerConfigurationV1` surface. Do not plan to treat `0.0`, `NaN`, or null as a substitute for that exception.
- Keep the class internal and preserve the block-scoped namespace and region organization used by other testers.

Must preserve:
- The common lifecycle and common-member tests inherited from `DeviceTesterBaseClass`.
- All existing COM behavior for other types.
- No device motion is initiated while the Aligner-specific methods remain placeholders.

Verification:
- Build the new file on both target frameworks.
- Inspect the source to confirm every IAligner property and method has exactly one visible TODO placeholder.
- Run an Alpaca Aligner conformance attempt and confirm it creates the client, validates interface version 1, connects, runs common tests, and exits without attempting the placeholder members.

### Step 4: Register the Aligner tester in the conformance manager

Location:
- `ConformU\ConformanceTestManager.cs`

Symbols:
- `ConformanceTestManager.AssignTestDevice()`

Changes:
1. Add a switch case for `DeviceTypes.Aligner` that constructs `AlignerTester(configuration, TL, cancellationToken, ConformCancellationTokenSource, conformResults)`.
2. Put the case near the alphabetically/currently ordered list, preferably at the start of the switch.
3. Keep the existing `default` invalid-device behavior unchanged.

Important details:
- This registration makes an Aligner selected through either GUI or CLI reach `AlignerTester.CreateDevice()`.
- Do not add a COM-specific check in `AssignTestDevice()`. `AlignerTester.CreateDevice()` owns the explicit Alpaca-only error.

Must preserve:
- Existing switch cases and tester instances for every current type.

Verification:
- Select an Alpaca Aligner URI through the CLI parsing path and confirm the run reaches `AlignerTester`.
- Attempt a COM Aligner selection and confirm the run fails with the Aligner-only technology message instead of attempting a facade or DriverAccess class.

### Step 5: Expose Aligner in the UI without exposing a COM path

Location:
- `ConformU\Pages\DeviceSelect.razor`

Symbols:
- `validAscomdDeviceTypes`
- `RadzenTabsItem Text="Windows COM Drivers"`
- `OnInitialized()`
- `OnDeviceTypeChange(object)`
- Alpaca discovery/manual-selection code

Changes:
1. Add `DeviceTypes.Aligner.ToString()` to `validAscomdDeviceTypes`. Keep the existing explicit string-list approach.
2. Modify the Windows COM Drivers tab's `Disabled` condition so it is disabled when either:
   - the OS is not Windows; or
   - `selectedDeviceTypeString == DeviceTypes.Aligner.ToString()`.
   Preserve the existing `Visible=@osIsWindows` behavior.
3. When a user changes the device-type dropdown to Aligner while the COM tab is selected, switch `selectedTabIndex` and `selectedTabNumber` to `ALPACA_TAB_NUMBER` and start/restart Alpaca discovery through the existing `StartDiscovery()` path.
4. In `OnInitialized()`, if the persisted device type is Aligner, do not select the COM tab even if stale settings say `DeviceTechnology.COM`; retain the Alpaca-tab index and use the existing Alpaca initialization path.
5. Leave the Alpaca discovery and manual selection code otherwise unchanged. Both already convert the selected name through `Devices.StringToDeviceType(...)` and already store `DeviceTypes.Aligner`.

Important details:
- A disabled COM tab is the UI guard. The conformance tester's explicit technology check remains the runtime guard for CLI/stale-settings callers.
- Do not silently delete or rewrite a persisted COM Aligner `ComDevice`; only prevent the stale selection from producing a selectable COM UI.
- Do not remove the existing Video-specific Alpaca restrictions.

Must preserve:
- All current COM selection behavior for non-Aligner device types.
- All current Alpaca selection, discovery, manual address, and navigation behavior.

Verification:
- On a non-Windows build, no COM UI changes are visible.
- On Windows, select every existing type and confirm the COM tab remains available exactly as before.
- Select Aligner and confirm the dropdown accepts it, the Alpaca controls work, and the COM tab cannot be selected.
- Manually configure an Alpaca Aligner and confirm settings contain `DeviceTechnology.Alpaca` and `DeviceType.Aligner`.

### Step 6: Register Aligner in the Alpaca protocol-test dispatch with placeholders

Location:
- `ConformU\AlpacaProtocolTestManager.cs`

Symbols:
- device-specific switch in `AlpacaProtocolTestManager`
- new `private async Task TestAligner()`

Changes:
1. Add `case DeviceTypes.Aligner` to the device-specific switch and `await TestAligner();`.
2. Add `TestAligner()` next to the other `Test<Type>` methods.
3. Follow existing setup/cleanup convention with an `AlpacaAligner` returned from `AlpacaClient.GetDevice<AlpacaAligner>(settings.AlpacaDevice, ...)`, `using` disposal, best-effort connect, `await TestConnect(aligner)`, and best-effort disconnect. Reuse the same `userAgentProductName`, `userAgentProductVersion`, username, and password arguments as adjacent test methods.
4. Add only comment placeholders for Aligner-specific protocol checks. Put `Configuration` first, then:
   - `IsMoving`;
   - `XPosition`;
   - `YPosition`;
   - `ZPosition`;
   - `FindHome`;
   - `Halt`;
   - `Move`;
   - `MoveAbsolute`;
   - `Sync`.
5. The `Configuration` placeholder must specifically state that the future protocol test should inspect the returned object and, for client-based verification, ensure the concrete runtime value implements `IAlignerConfigurationV1` and exposes all 21 configuration properties. Do not implement those requests or JSON assertions now.
6. Do not add a new `AdditionalParameterCheck` enum value in this phase.

Important details:
- The raw protocol manager still tests common Alpaca behavior through its existing common-test code. The new method only supplies type routing and the placeholder section.
- Keep `TestCancelledException` handling and disconnect cleanup consistent with adjacent methods.
- Do not duplicate the conformance tester's detailed property assertions here; protocol checks remain raw-endpoint checks.

Must preserve:
- Existing common protocol parameter, transaction ID, status, casing, connection, DeviceState, and disconnect tests.
- Existing type-specific protocol methods for every other device.

Verification:
- Run the Alpaca protocol command against an Aligner endpoint. It must no longer throw `Unknown device type`.
- The protocol run must exercise existing common checks, enter `TestAligner()`, execute no Aligner-specific placeholder request, and disconnect cleanly.

### Step 7: Add focused automated coverage for the capability shim

Location:
- Add `ConformU.Tests\DeviceCapabilitySupportTests.cs`.

Symbols:
- `DeviceCapabilitySupportTests`

Changes:
1. Add xUnit tests for the wrapper. Because the production assembly is strongly signed and the helper is intentionally internal, access it through reflection from `typeof(ConformConfiguration).Assembly`. Resolve `DeviceCapabilitySupport` with `Assembly.GetType("ConformU.DeviceCapabilitySupport", throwOnError: true)`, and get each `BindingFlags.NonPublic | BindingFlags.Static` method by name. Invoke the methods with typed arguments. Do not add `InternalsVisibleTo` or sign the test assembly.
2. Test:
   - Aligner V1 returns `true` from both `HasConnectAndDeviceState` and `IsPlatform7OrLater`.
   - Aligner versions 0 and 2 return `false` from both methods.
   - A representative existing device, for example `Camera` V4, returns the same booleans as direct calls to the external `DeviceCapabilities` methods.
   - `TryGetLatestPlatform6Interface(DeviceTypes.Aligner, out _)` returns `false` and a zero out value.
   - For an existing type known to exist in `LatestPlatform6Interface`, the helper returns `true` and the same value as the Library dictionary.
3. For the `out short` helper, pass an initially zero-valued `short` in an `object[]`, then assert the returned `bool` and the updated array element.
4. Do not write unit tests that claim to test the placeholder-specific Aligner properties or methods. There is no Aligner endpoint fixture and the requested implementation phase does not implement those tests.

Important details:
- The test class verifies only ConformU's compatibility logic. It must not assert unrelated Library behavior beyond the one representative delegation comparison.
- Keep the compatibility class internal. Do not add a public API, `InternalsVisibleTo`, a signing change, or a package reference just to access internals.

Must preserve:
- Existing `ExtensionMethodsTests` and xUnit test output.

Verification:
- Run the in-process xUnit runner as specified in the Tests section. Expected: all existing two tests plus the new focused tests pass.

### Step 8: Perform complete build and manual Aligner-route verification

Location:
- No additional file changes.

Symbols:
- Entire solution and both runtime paths.

Changes:
1. Build the full solution for both target frameworks.
2. Run the xUnit test host.
3. Perform the manual conformance and protocol checks described below when an actual Aligner device or simulator is available.
4. Review the final diff to ensure there is no `AlignerFacade`, no `DriverAccess.Aligner` instantiation, no COM branch in `AlignerTester`, and no implemented Aligner-specific property/method body.

Important details:
- The repository does not contain an Aligner simulator, so this implementation cannot fully exercise a live `Configuration` response. A placeholder/source-order review is required even if live integration is unavailable.
- Do not treat a successful build as proof that `AlpacaAligner.Configuration` can deserialize against a real endpoint. That is an unresolved Library-level integration check.

Must preserve:
- Clean behavior for all existing device types.
- A zero issue/error result for unrelated common test cases.

Verification:
- See the complete command list and expected outcomes in `Tests`.

## Tests

### Existing automated tests that must continue to pass

Run:

```text
dotnet run --project ConformU.Tests\ConformU.Tests.csproj -c Debug -- -noLogo -noColor
```

The existing `ExtensionMethodsTests.ToJsonNameCaseSensitivity_WhenTrue_ReturnsCorrectCasingOnly` and `ExtensionMethodsTests.ToJsonNameCaseSensitivity_WhenFalse_ReturnsAnyCasing` must still pass.

Do not use plain `dotnet test ConformU.sln` as the verification command unless the test-project MTP configuration is separately fixed. The current repository reports that VSTest targeting is no longer supported by Microsoft.Testing.Platform on the installed .NET 10 SDK. Changing that test infrastructure is outside this task.

### New automated tests

Add the `DeviceCapabilitySupportTests` cases listed in Step 7. Expected outcomes:

- Aligner V1 is treated as supporting `Connect`/`Disconnect`/`Connecting`/`DeviceState` and as a Platform 7 interface.
- Aligner interface versions 0 and 2 are not treated as valid by the wrapper.
- Existing device behavior remains identical to direct Library calls.
- Aligner has no Platform 6 interface value in the wrapper, so error reporting cannot trigger a dictionary lookup failure.

### Build and static verification

Run:

```text
dotnet build ConformU.sln -c Debug -p:Platform="Any CPU"
```

Expected: both `net10.0` and `net10.0-windows` compile with zero errors. Do not accept a one-target-only build.

Then verify:

1. `ConformU\Conform\AlignerTester.cs` imports no COM namespace and has no `#if WINDOWS` COM branch.
2. `CheckProperties()` starts with the `Configuration` placeholders and places `IsMoving`, `XPosition`, `YPosition`, and `ZPosition` afterward.
3. Every member of `IAlignerV1` and every member of `IAlignerConfigurationV1` has a TODO placeholder.
4. `CheckMethods()` contains comments only for `FindHome`, `Halt`, `Move`, `MoveAbsolute`, and `Sync`.
5. `DeviceTesterBaseClass` contains the four expected Aligner `DeviceState` property names.
6. No remaining direct `DeviceCapabilities.HasConnectAndDeviceState` or `IsPlatform7OrLater` capability call occurs in the modified lifecycle/protocol code.
7. No source-code change references `AlignerFacade`, `DriverAccess.Aligner`, or an Aligner COM ProgID.

### Manual positive tests

1. Start the GUI and select `Aligner`:
   - expected: Aligner is present in the device-type dropdown;
   - expected on Windows: COM Drivers is disabled;
   - expected: Alpaca discovery and manual configuration remain usable.
2. Select or discover an Alpaca Aligner and start a conformance run:
   - expected: `AlignerTester` is selected;
   - expected: an `AlpacaAligner` instance is created;
   - expected: interface version 1 validates;
   - expected: common connect/common-member/DeviceState/disconnect paths run;
   - expected: no `KeyNotFoundException` names Aligner;
   - expected: placeholder Aligner-specific checks do not issue movement requests.
3. Run the CLI conformance command with an Alpaca Aligner URI:
   - expected: the URI is accepted by the dynamic `Devices.DeviceTypeNames()` parser and runs through `AlignerTester`.
4. Run the Alpaca protocol checker for the same device:
   - expected: no `Unknown device type` error;
   - expected: common protocol tests execute;
   - expected: `TestAligner()` is reached and exits through its placeholders and normal cleanup.
5. With a real Aligner endpoint that returns Configuration data, manually inspect the returned `Configuration` in a debugger or a temporary diagnostic:
   - expected: it is non-null;
   - expected: the runtime type is assignable to `IAlignerConfigurationV1`;
   - expected: all 21 configuration properties are readable or, for the documented optional numeric members, throw `PropertyNotImplementedException` consistently with the interface.
   Do not convert this diagnostic into implemented conformance-test code in this phase.

### Negative and regression cases

1. Attempt a COM or stale persisted COM Aligner:
   - expected: `AlignerTester.CreateDevice()` produces an explicit Alpaca-only `InvalidValueException`;
   - expected: no COM facade or DriverAccess code is invoked.
2. Force an Aligner to report interface version 2:
   - expected: `ValidateInterfaceVersion()` rejects it through existing Library validation;
   - expected: error reporting does not throw when it checks for a missing Platform 6 latest-version entry.
3. Exercise Camera, CoverCalibrator, Dome, FilterWheel, Focuser, ObservingConditions, Rotator, SafetyMonitor, Switch, Telescope, and Video normal dispatch:
   - expected: all existing tester assignments remain unchanged;
   - expected: non-Aligner capability results are delegated to the Library exactly as before.
4. Disable property or method testing in settings:
   - expected: `AlignerTester.CheckConfiguration()` records the same configuration alerts as the existing simple testers.

## Do Not Change

- Do not change package versions in `ConformU\ConformU.csproj`, including the 4.2.0-rc.1 ASCOM packages.
- Do not add an Aligner COM facade, `DriverAccess` wrapper, `DriverHostForm` use, `#if WINDOWS` branch in `AlignerTester`, or Aligner registration for COM.
- Do not implement the Aligner property or method test bodies. This phase adds ordered placeholders only.
- Do not add an Aligner settings class, persisted `Settings` properties, or increment `Settings.SETTINGS_COMPATIBILTY_VERSION`.
- Do not alter the existing test execution order, connection/disconnection behavior, or test results for non-Aligner devices.
- Do not change `Program.GetDeviceTechnology()`; it already uses the Library's current device-type names, which include Aligner.
- Do not modify `newalignerdevice.md`, generated build output, publish scripts, or unrelated formatting.
- Do not add a ConformU replacement implementation of `IAlignerConfigurationV1` or deserialize a custom DTO in application code.
- Do not change the existing xUnit/Microsoft.Testing.Platform project configuration solely to make `dotnet test` work.
- Do not make the internal capability wrapper public or add friend-assembly/signing changes.

## Risks and Implementation Traps

- **Calling the Library capability helpers directly for Aligner:** `HasConnectAndDeviceState`, `IsPlatform7OrLater`, and the Platform 6 dictionary lookups currently throw or omit Aligner. Route all affected calls through `DeviceCapabilitySupport`.
- **Treating Aligner V1 as a historic synchronous interface:** Aligner V1 inherits `IAscomDeviceV2` and must use the Connect/Disconnect/Connecting/DeviceState lifecycle. The wrapper intentionally treats valid Aligner V1 as Platform 7.
- **Accidentally implementing COM because most existing testers have COM branches:** `AlignerTester` must be Alpaca-only. A copied tester is likely to include COM `using` directives or a facade branch; remove those immediately.
- **Implementing the requested placeholders:** The task requires comments, not executable test logic. Calling `Move`, `Halt`, `FindHome`, `MoveAbsolute`, or `Sync` would violate scope and could move hardware.
- **Testing `Configuration` after the other properties:** The first IAligner-specific test placeholder must be Configuration. Future implementations need its capability flags before selecting method/property expectations.
- **Trusting the compile-time `IAlignerConfigurationV1` return type:** The contract must be checked against the concrete runtime type. Current binary reflection reports that `ASCOM.Alpaca.Clients.AlignerConfigurationV1Dto` has no public properties and implements no interfaces, even though XML documentation describes its properties. A future Configuration test must expose this mismatch rather than assume the cast is valid.
- **Adding `Configuration` to DeviceState:** The Alpaca `AlignerState` model exposes `IsMoving`, `XPosition`, `YPosition`, and `ZPosition`, not Configuration. Adding Configuration as an operational-state member would incorrectly flag valid DeviceState responses.
- **Breaking the UI by listing Aligner for COM:** The current selector shares one list. Add Aligner to the list, but explicitly disable/reroute the COM tab when it is selected.
- **Adding a protocol method that returns without cleanup:** Follow the adjacent `Test<Type>` pattern and preserve cancellation/disconnect handling even though the type-specific assertions are placeholders.
- **Changing global capability behavior while fixing the missing dictionary entries:** The support class must delegate all non-Aligner inputs unchanged. Do not catch exceptions globally or silently return defaults for other types.

## Assumptions / Unresolved Questions

- The implementation assumes Aligner interface version 1 is the only valid version, as reported by the referenced Library's `LatestInterface` and `IsValidAscomInterface` APIs.
- No Aligner simulator or live endpoint is present in the repository. Runtime behavior of `AlpacaAligner.Configuration` against a real JSON response cannot be verified during planning.
- A significant Library inconsistency remains unresolved: `AlpacaAligner.Configuration` is declared to return `IAlignerConfigurationV1` and its getter calls `RemoteDevice.GetValue<IAlignerConfigurationV1>()`, but reflection of the referenced 4.2.0-rc.1 binary shows that the documented `AlignerConfigurationV1Dto` implements no interfaces and exposes no public properties. The application implementation must not work around this. The future Configuration test placeholder is the place to detect and report a failure. If live integration fails before a usable object is returned, that is a Library/package issue requiring correction upstream, not a reason to add an ad-hoc ConformU DTO.
- The interface documentation supplied by the Library does not justify imposing additional numeric ranges on optional configuration values. Future tests should use only the documented mandatory/optional and paired-property rules.

## Implementation Checklist

1. Add `DeviceCapabilitySupport.cs` with exact Aligner V1 capability exceptions and guarded Platform 6 lookup.
2. Replace all relevant shared and protocol capability calls with the wrapper.
3. Add Aligner to the `DeviceState` expected-property switch.
4. Make `ConformanceTestManager` error reporting safe for a type with no Platform 6 dictionary entry.
5. Add `Conform\AlignerTester.cs` with an Alpaca-only `CreateDevice`.
6. Add all IAligner and configuration TODO placeholders, with Configuration first and no actual device-specific test code.
7. Register `DeviceTypes.Aligner` in `ConformanceTestManager.AssignTestDevice()`.
8. Add Aligner to the GUI list and disable/reroute the COM tab.
9. Add Aligner protocol dispatch and a `TestAligner()` placeholder method.
10. Add focused capability-wrapper unit tests without changing production visibility or signing.
11. Build both target frameworks and run the xUnit v3 in-process test host.
12. Review the source/diff for COM code, implemented Aligner tests, placeholder ordering, and complete DeviceState coverage.
13. When an Aligner endpoint is available, perform the positive, COM-negative, and invalid-interface-version integration checks.
