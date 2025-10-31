# Core Service & Model Inventory

The following classes originated under `EpicPlanner.Core/Services` and `EpicPlanner.Core/Model` and are consumed by both `EpicPlanner.Planner` and `EpicPlanner.Checker`. The table highlights the shared responsibilities alongside planner-only and checker-only behaviors so the split between the applications remains explicit.

| Class | Shared responsibility | Planner-specific behavior | Checker-specific behavior |
| --- | --- | --- | --- |
| `BusinessCalendar` | Provides working-day calculations that normalize resource capacity for holidays and weekends. | Used indirectly through the planner data provider when preparing sprint capacities. | The checker reuses the same capacity adjustments so its reports align with the planner's calculations. |
| `PlanningDataProvider` | Loads epics, capacities, and Redmine data into shared snapshot components. | Produces `PlannerPlanningSnapshot` instances without checker-only enrichment when analysis mode skips planned hours. | Builds `CheckerPlanningSnapshot` instances enriched with planned hours and epic summaries to validate actual allocations. |
| `PlannerPlanningSnapshot` / `CheckerPlanningSnapshot` | Specializations that transform snapshot components into simulator-ready structures via the shared `PlanningSnapshotBase`. | `PlannerPlanningSnapshot` creates filtered planner simulators for analysis mode before exporting planning outputs. | `CheckerPlanningSnapshot` supplies checker simulators with full data for comparison and epic-state checks. |
| `SimulatorBase` | Provides the shared scheduling engine and metrics collection reused by both applications. | `PlannerSimulator` extends the base to export planning Excel files and sprint-based Gantt charts, including analysis-specific date alignment. | `CheckerSimulator` extends the base to produce comparison and epic state reports that surface discrepancies. |
| `Epic` | Represents product epics, including dependencies and assignment wishes. | Analysis mode walks dependency graphs to determine scope and timeline alignment. | Checker reports evaluate epic states, remaining work, and planned capacity discrepancies. |
| `Allocation` | Tracks historical allocation slices for reporting. | Included in exported planning spreadsheets and Gantt charts. | Surfaces in checker output to compare planned versus consumed allocations. |
| `ResourceCapacity` | Encapsulates development, maintenance, and analysis capacities. | Adjusted per sprint to inform planning simulations. | The same adjusted capacities drive checker validations for accuracy. |
| `SprintEpicSummary` | Stores consumption metrics per epic and sprint. | Displayed in planning exports alongside capacity allocations. | Compared with Redmine data in checker reports for variance analysis. |
| `Wish` | Captures desired resource allocations parsed from the backlog. | Feeds the simulator's planning algorithm when scheduling work. | Checker output cites wish data when highlighting allocation gaps. |

Classes that are application-specific are now scoped under `EpicPlanner.Core.Planner` (`PlanningRunner`, `EnumPlanningMode`, planner DI extensions) and `EpicPlanner.Core.Checker` (`CheckingRunner`, `EnumCheckerMode`, checker DI extensions`).
