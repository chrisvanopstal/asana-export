$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols

# Script variables (replace with your asana IDs)
$asanaToken = "Bearer [Your Asana API key]" # Example: "Bearer 0/234de3f8d34..." 
$workspaceId = "[Your Asana workspace ID]" # Example: 6285880...

# The directory to save exported data
$filePath = "C:\[Your export file path]" # Example: C:\Export

# Loop through projects
$tasksArray = $tasksToHeartsArray = $tasksToTagsArray = $tasksToFollowersArray = $tasksToMembershipsArray = $tasksToStoriesArray = @()
$projects = Invoke-WebRequest -Uri "https://app.asana.com/api/1.0/workspaces/$($workspaceId)/projects" -Headers @{"Authorization"=$asanaToken} | ConvertFrom-Json | Select -expand data
foreach ($project in $projects) {

    # Get tasks for each project
    echo "Getting tasks for $($project.name)"
    $url = "https://app.asana.com/api/1.0/projects/" + $project.id + "/tasks"

    $tasks = Invoke-WebRequest -Uri $url -Headers @{"Authorization"=$asanaToken} | ConvertFrom-Json | Select -expand data
    echo "$($tasks.length) tasks found"

    # Get task details
    foreach ($task in $tasks) {

        # Ensure task doesn't already exist (in case the task exists in multiple projects)
        if (($tasksArray | Where-Object { $_.id -eq $task.id }).length -eq 0) {

            $url = "https://app.asana.com/api/1.0/tasks/" + $task.id
            $taskDetails = Invoke-WebRequest -Uri $url -Headers @{"Authorization"=$asanaToken} | ConvertFrom-Json | Select -expand data
            $tasksArray += $taskDetails

            # Pull child objects into their own arrays so we can save down to CSV files
            $taskDetails.memberships | Add-Member "task_id" $task.id
            foreach ($taskMembership in $taskDetails.memberships) {
                $tasksToMembershipsArray += $taskMembership
            }
            
            $taskDetails.hearts | Add-Member "task_id" $task.id
            foreach ($taskHeart in $taskDetails.hearts) {
                $tasksToHeartsArray += $taskHeart
            }

            $taskDetails.followers | Add-Member "task_id" $task.id
            foreach ($taskFolower in $taskDetails.followers) {
                $tasksToFollowersArray += $taskFolower
            }

            $taskDetails.tags | Add-Member "task_id" $task.id
            foreach ($taskTags in $taskDetails.tags) {
                $tasksToTagsArray += $taskTags
            }

            # Retrieve stories
            $url = "https://app.asana.com/api/1.0/tasks/" + $task.id + "/stories"
            $taskStories = Invoke-WebRequest -Uri $url -Headers @{"Authorization"=$asanaToken} | ConvertFrom-Json | Select -expand data
            foreach ($story in $taskStories) {
                $story | add-member task_id $task.id
                $tasksToStoriesArray += $story
            }

            # Pull out first story so we know who created the task (useful for reporting)
            $taskDetails | add-member created_by ($taskStories | Select-object -first 1 | Select -expand created_by)
        }
    }
}

# Write arrays to CSV files
echo "Writing files to: $($filePath)"
$tasksArray | Select @{name="task_id";expression={$_.id}}, created_at, @{name="created_by_id";expression={$_.created_by.id}}, @{name="created_by_name";expression={$_.created_by.name}}, modified_at, name, @{name="type";expression={$_.custom_fields[0].enum_value.name}}, @{name="source";expression={$_.custom_fields[1].enum_value.name}}, completed, assignee_status, completed_at, due_on, due_at, num_hearts, parent, hearted, @{name="workspace_id";expression={$_.workspace.id}}, @{name="workspace_name";expression={$_.workspace.name}}, @{name="assignee_id";expression={$_.assignee.id}}, @{name="assignee_name";expression={$_.assignee.name}} | Export-Csv -path "$($filePath)\tasks.csv" -notype
$tasksToMembershipsArray | Select task_id, @{name="project_id";expression={$_.project.id}}, @{name="project_name";expression={$_.project.name}}, @{name="section_id";expression={$_.section.id}}, @{name="section_name";expression={$_.section.name}} | Export-Csv -path "$($filePath)\memberships.csv" -notype
$tasksToTagsArray | Select task_id, @{name="tag_id";expression={$_.id}}, @{name="tag_name";expression={$_.name}} | Export-Csv -path "$($filePath)\tags.csv" -notype
$tasksToFollowersArray |  Select task_id, @{name="follower_id";expression={$_.id}}, @{name="follower_name";expression={$_.name}} | Export-Csv -path "$($filePath)\followers.csv" -notype
$tasksToStoriesArray | Select task_id, @{name="story_id";expression={$_.id}}, created_at, type, text, @{name="created_by_id";expression={$_.created_by.id}}, @{name="created_by_name";expression={$_.created_by.name}} | Export-Csv -path "$($filePath)\stories.csv" -notype