# BP System Refactoring Specifications

## Canonical Model
This section should outline the canonical model for the BP system, detailing the structure and references for each component involved in the refactoring process.

## Storage Decisions
- Define the decisions on data storage before and after refactoring.
- Explain the transitions between previous and new storage methodologies, including any technology stack changes.

## Manual Adjustments
- List any manual adjustments necessary during the refactoring process.
- Specify procedures for modifying existing data to comply with the new model.

## Attendance Scanning
- Explain how attendance scanning will be integrated into the new BP system.
- Provide detailed use cases and any technological or logistical considerations that must be taken into account.

## Pull Request Requirements
1. **Specification Documentation**: Each pull request must include documentation outlining the changes and their purpose.
2. **Code Review**: Pull requests must undergo a code review process with at least two reviewers.
3. **Testing**: All new features must have associated unit tests that demonstrate functionality.
4. **Deployment Instructions**: Clear instructions must be provided for deploying changes in the production environment.
5. **Backward Compatibility**: Ensure that changes made in the pull request do not break existing functionality for current users.