| Relationship Type | Description or Purpose |
|-------------------|------------------------|
| Depends on; Used by | This indicates that a CI depends on another to function (App → Database). |
| Runs on; Hosts | This indicates that a CI runs on another CI (App → Server or VM → Host). |
| Connected to | This is a generic connection, often used for networking hardware. |
| Contained in; Contains | This indicates that a CI is part of a larger CI (VM → Cluster, File → Folder). |
| Backs up; Backed up by | This defines data backup or recovery relationship. |
| Monitored by; Monitors | This indicates a CI is being monitored (App → Monitoring Tool). |
| Installed on; Installs | This shows deployment (Software → Machine). |
| Owned by; Owns | This links CI to an owner, such as a person, team, or department. |
| Connected via | This describes a CI connected through a specific path or medium, which could be a cable. |
| Part of; Has Part | This is used in hardware to show a component structure (Disk → Server). |
| Impacts; Impacted by | This defines service dependencies for impact analysis. |
| Requires; Required by | This is similar to depends on, used in service design. |
| Relates to | This is a catch-all for non-hierarchical associations. |
| Accessed by; Accesses | This shows user or system access to a CI (User → System). |
| Deployed on | This indicates that a CI is deployed on another CI (Docker container → Kubernetes node). |
| Supports; Supported by | This indicates support structure or fallback (Primary server → Secondary server). |
| Replaces; Replaced by | The CIs have been retired and replaced by a new CI. |