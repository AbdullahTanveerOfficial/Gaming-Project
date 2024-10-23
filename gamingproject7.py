import pandas as pd

# Define a class for Players
class Player:
    def __init__(self, PlayerID, Handicap, preferencias):
        self.PlayerID = PlayerID
        self.Handicap = float(Handicap)  # Convert to float
        self.preferred_players = [preferencias[f'Preferencia{i}'] for i in range(1, 9) if pd.notna(preferencias[f'Preferencia{i}'])]

    def __repr__(self):
        return f"Player({self.PlayerID}, Handicap: {self.Handicap}, preferred: {self.preferred_players})"


# Function to load players from an Excel file
def load_players_from_excel(file_path):
    """
    Load players from an Excel file.
    Excel format: PlayerID, Handicap, Preferencia1 to Preferencia8
    """
    players = pd.read_csv(file_path, nrows=None)
    players_list = players.to_dict(orient='records')
    player_objects = [Player(str(p['PlayerID']), p['Handicap'], p) for p in players_list]
    
    return player_objects


# Function to create groups based on preferences only
def create_groups_based_on_preferences(players_list):
    """
    Create groups purely based on mutual preferences.
    
    :param players_list: List of Player objects.
    :return: List of groups and list of ungrouped players.
    """
    groups = []
    ungrouped_players = players_list.copy()

    while ungrouped_players:
        requesting_player = ungrouped_players.pop(0)
        group = [requesting_player]
        
        # Find mutual preferences for this player
        for preferred_player_id in requesting_player.preferred_players:
            for player in ungrouped_players:
                if player.PlayerID == preferred_player_id and requesting_player.PlayerID in player.preferred_players:
                    group.append(player)
                    ungrouped_players.remove(player)
                    
                    if len(group) == 4:  # If the group reaches size 4, stop adding
                        break
            if len(group) == 4:
                break

        # Add the group to the list
        groups.append(group)
    
    return groups, ungrouped_players


# Function to merge leftover players into existing groups
def merge_remaining_players(groups):
    """
    Merge remaining players into existing groups, ensuring no group exceeds 4 players.
    Groups of 2 are allowed, but no group should have more than 4 players.
    :param groups: List of groups formed based on preferences.
    :return: Final list of merged groups.
    """
    final_groups = []
    remaining_players = []

    # Collect groups with less than 3 players and build final groups with 3 or 4 players
    for group in groups:
        if len(group) < 3:
            remaining_players.extend(group)
        else:
            final_groups.append(group)

    # Merge remaining players
    while remaining_players:
        if len(remaining_players) == 1:
            # Add the last remaining player to a group of 3 (if exists)
            added = False
            for group in final_groups:
                if len(group) == 3:
                    group.append(remaining_players.pop(0))
                    added = True
                    break
            # If we couldn't add the player to any group, create a new group
            if not added:
                final_groups.append(remaining_players)
                remaining_players = []
        elif len(remaining_players) == 2:
            # Form a group of 2 if we can't merge them into another group
            final_groups.append(remaining_players[:2])
            remaining_players = []
        elif len(remaining_players) >= 3:
            # Create a new group of 3 or 4
            group_size = 4 if len(remaining_players) >= 4 else 3
            new_group = remaining_players[:group_size]
            remaining_players = remaining_players[group_size:]
            final_groups.append(new_group)

    group_of_2 = []
    for f_group in final_groups:
        if len(f_group) == 2:
            group_of_2.append(f_group)
            final_groups.remove(f_group)

    while len(group_of_2) > 1:
        g1 = group_of_2.pop()
        g2 = group_of_2.pop()
        final_groups.append(g1 + g2)

    if len(group_of_2) == 1:
        for c_group in final_groups:
            if len(c_group) == 4:
                rem_player = c_group.pop()
                group_of_2[0].append(rem_player)
                final_groups.append(group_of_2[0])
                break

    return final_groups


# Function to save the final groups to an Excel file
def save_groups_to_excel(groups, file_name="groups.xlsx"):
    """
    Save the created groups to an Excel file, sorted by the average handicap in ascending order.
    Format of the Excel file:
    Group N | Player | Handicap | Avg Handicap Group
    """
    group_data = []
    # Calculate the average handicap for each group and prepare group data
    groups_with_avg = [(group, sum([player.Handicap for player in group]) / len(group)) for group in groups]

    # Sort the groups based on the average handicap in ascending order
    sorted_groups = sorted(groups_with_avg, key=lambda x: x[1])

    for idx, (group, avg_handicap) in enumerate(sorted_groups):
        for i, player in enumerate(group):
            if i == 0:
                group_data.append([idx + 1, player.PlayerID, player.Handicap, round(avg_handicap, 2)])
            else:
                group_data.append([None, player.PlayerID, player.Handicap, None])

    df_groups = pd.DataFrame(group_data, columns=["Group N", "Player", "Handicap", "Avg Handicap Group"])
    
    # Save the sorted groups to an Excel file
    df_groups.to_excel(file_name, index=False)

    print(f"Groups saved to {file_name}, sorted by average handicap.")



# Main function to run the process
if __name__ == "__main__":
    file_path = "INPUT_EXAMPLE_2_SMALL.csv"  # Replace with actual file path
    
    # Step 1: Load players from Excel
    players_list = load_players_from_excel(file_path)
    
    # Step 2: Create groups based on preferences only
    preference_groups, ungrouped_players = create_groups_based_on_preferences(players_list)
    
    # Step 3: Merge remaining players into groups
    final_groups = merge_remaining_players(preference_groups)
    
    # Step 4: Save the groups to an Excel file
    save_groups_to_excel(final_groups, "final_groups.xlsx")

    print(f"Groups have been created and saved to 'final_groups.xlsx'.")
