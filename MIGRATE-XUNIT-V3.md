MIGRATION: xUnit v2 -> xUnit v3

Rationale
- NuGet and repository maintainers recommend migrating to xUnit v3 when feasible. xUnit v3 introduces breaking changes but also improvements.

This file records the repository-level decision to migrate test projects to xUnit v3 in future.

Status
- As of this change, xUnit v3 is not available as a stable NuGet package for this solution (latest stable xunit: 2.9.3).
- We add this note to make intent explicit for future GitHub Copilot agents and contributors.

Next steps
1. When xUnit v3 stable is published to nuget.org, update HVM Kasserer.Tests\HVM Kasserer.Tests.csproj:
   - Replace <PackageReference Include="xunit" Version="2.9.3" /> with the xUnit v3 package version.
   - Replace or update any runner/adapter packages as required (xunit.runner.visualstudio or new adapters).
2. Build and fix compiler errors in tests (Assert API, attributes, test helpers may change).
3. Run tests and fix behavioral changes.
4. Update CI to use matching runners/adapters if CI exists.

Notes
- This repository change was applied directly on the current branch per request.
- See the commit that added this file for details.
