% === By Nathan Au August 2025 ===

%% SCRIPT: GENERATE ANIMATED HEAT-MAP (WITH LATE/EARLY LOGIC)
clc; clear; close all; 
disp("Starting script...");

% === File Setup ===
data_file = 'Animation_Data.xlsx'; 
arrival_sheet = 'Arrival';
departure_sheet = 'Departure';
map_file = 'Aerodrome Map.xlsx';

% === Parameters ===
num_intervals = 288; % 24 hours * 12 (5-min intervals per hour)
num_days = 31;       % Divisor for calculating daily average

% === Read Airport Map Layout ===
[~, ~, raw_map] = xlsread(map_file);
disp("Map layout loaded.");

% === Initialize Time-based Mechanics Map ===
mechanic_counts_over_time = cell(num_intervals, 1);
for i = 1:num_intervals
    mechanic_counts_over_time{i} = containers.Map('KeyType', 'char', 'ValueType', 'any');
end

% === Define Airline Groupings ===
rule_keys = {'CI','MU','CZ','CA','3U','FM','MF','SC','ZH','HB','6E','9C','AI','BA','BX','HO','HU','LH','LJ','MS','UQ','NS','O3','OD','OZ','QW','RA','TV','HX','KR','OM','WW','UO'};
rule_group_values = [1,2,2,2,2,2,2,2,2,3,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,4,5,5,5,5,6];
rule_group_map = containers.Map(rule_keys, rule_group_values);
report_keys = {'UO','HX','HB','MU','CA','ZH','CZ','HU','SC','FM','MF','O3','3U','HO','TV','9C','CK','NS','UQ','PN', 'CI', 'OZ','AI','MTJ','OM','6E','OD','BX','LJ','KR','MS','RA','B3', 'LH','BA','WW','QW','KJ', '5H', 'AA', 'IT'};
report_group_values = [1, 3, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2, 3, 3, 3, 3];
report_group_map = containers.Map(report_keys, report_group_values);

% =========================================================================
%% === MODIFIED SECTION: Process Both Arrival and Departure Sheets ===
% This section now uses the advanced "late vs. early" logic.
% =========================================================================
disp("Starting data processing with late/early logic...");

for pass = 1:2
    if pass == 1; sheet_name = arrival_sheet; is_departure = false; else; sheet_name = departure_sheet; is_departure = true; end
    fprintf('Processing rules from sheet: %s...\n', sheet_name);
    
    opts = detectImportOptions(data_file, 'Sheet', sheet_name);
    opts.DataRange = 'A2';
    % Read columns 3 (Sched), 9 (Actual), 16 (Bay), 17 (Airline)
    opts.SelectedVariableNames = [3, 9, 16, 17]; 
    data = readtable(data_file, opts);

    for j = 1:height(data)
        scheduled_entry = string(data{j,1});
        actual_entry = string(data{j,2});
        bay = string(data{j,3});
        airline_code = upper(strtrim(string(data{j,4})));

        if ismissing(scheduled_entry) || ismissing(actual_entry) || ismissing(bay) || ismissing(airline_code)
            continue;
        end
        if ~isKey(rule_group_map, airline_code) || ~isKey(report_group_map, airline_code)
            continue;
        end

        try
            [scheduled_time, scheduled_time_in_minutes] = parse_time(scheduled_entry);
            [actual_time, actual_time_in_minutes] = parse_time(actual_entry);
        catch
            continue;
        end
        
        rule_group_num = rule_group_map(airline_code);
        report_group_num = report_group_map(airline_code);
        bay_name = upper(strtrim(bay));
        if endsWith(bay_name, "L") || endsWith(bay_name, "R")
            bay_name = extractBefore(bay_name, strlength(bay_name));
        end

        times = {}; personnel = [];
        if ~is_departure
            switch rule_group_num; case 1; times = {[-20, 20], [20, 35]}; personnel = [2, 1]; case 2; times = {[-25, 35]}; personnel = [1]; case 3; times = {[-20, 35]}; personnel = [1]; case 4; times = {[-20, 35]}; personnel = [1]; case 5; times = {[-20, 35]}; personnel = [1]; case 6; times = {[-20, 25]}; personnel = [1]; end
        else
            switch rule_group_num; case 1; times = {[-70, -25], [-25, 15]}; personnel = [1, 2]; case 2; times = {[-70, -15], [-15, 15]}; personnel = [1, 2]; case 3; times = {[-70, -20], [-20, 15]}; personnel = [1, 2]; case 4; times = {[-70, -15], [-15, 15]}; personnel = [1, 2]; case 5; times = {[-70, -15], [-15, 15]}; personnel = [1, 2]; case 6; times = {[-70, -20], [-20, 15]}; personnel = [1, 2]; end
        end
        
        if actual_time > scheduled_time % SCENARIO 1: FLIGHT IS LATE
            time_diff_minutes = minutes(actual_time - scheduled_time);
            delay_extension = 5 * ceil(time_diff_minutes / 5);
            if rule_group_num == 1 && ~is_departure % Special Case: Group 1 Arrival
                mechanic_counts_over_time = apply_rule(mechanic_counts_over_time, bay_name, report_group_num, scheduled_time_in_minutes, times{1}, personnel(1), num_intervals);
                ext_start_min = scheduled_time_in_minutes + times{1}(2);
                ext_end_min = ext_start_min + delay_extension;
                mechanic_counts_over_time = apply_rule(mechanic_counts_over_time, bay_name, report_group_num, 0, [ext_start_min, ext_end_min], personnel(1), num_intervals);
                shift = times{2}(2) - times{2}(1);
                shifted_start_min = ext_end_min;
                shifted_end_min = shifted_start_min + shift;
                mechanic_counts_over_time = apply_rule(mechanic_counts_over_time, bay_name, report_group_num, 0, [shifted_start_min, shifted_end_min], personnel(2), num_intervals);
            else % General Rule: All other late flights
                for k = 1:length(times)
                    mechanic_counts_over_time = apply_rule(mechanic_counts_over_time, bay_name, report_group_num, scheduled_time_in_minutes, times{k}, personnel(k), num_intervals);
                end
                last_rule_end_min = scheduled_time_in_minutes + times{end}(2);
                ext_start_min = last_rule_end_min;
                ext_end_min = ext_start_min + delay_extension;
                mechanic_counts_over_time = apply_rule(mechanic_counts_over_time, bay_name, report_group_num, 0, [ext_start_min, ext_end_min], personnel(end), num_intervals);
            end
        else % SCENARIO 2: FLIGHT IS EARLY or ON TIME
            rounded_down_minutes = 5 * floor(actual_time_in_minutes / 5);
            rounded_up_minutes   = 5 * ceil(actual_time_in_minutes / 5);
            for k = 1:length(times)
                start_min = rounded_down_minutes + times{k}(1);
                end_min = rounded_up_minutes + times{k}(2);
                mechanic_counts_over_time = apply_rule(mechanic_counts_over_time, bay_name, report_group_num, 0, [start_min, end_min], personnel(k), num_intervals);
            end
        end
    end
end
fprintf("✔ Data processing complete.\n");

%% === Create Animation ===
disp("Generating animation with group colors...");
video_writer = VideoWriter('mechanic_heatmap.mp4', 'MPEG-4');
video_writer.FrameRate = 2;
open(video_writer);

% --- Create Figure and Main Axes ---
fig = figure('Name', 'Animated Mechanic Heatmap', 'Position', [100, 100, 1200, 800]); % Made figure wider
main_ax = axes('Position', [0.05 0.1 0.7 0.85]); % Main axes for the heatmap

% --- Colormaps and Global Max Calculation ---
light_blue = [0.8, 0.9, 1.0]; dark_blue = [0.0, 0.2, 0.5];
light_red = [1.0, 0.8, 0.8];  dark_red = [0.6, 0.0, 0.0];
light_green = [0.8, 1.0, 0.8]; dark_green = [0.0, 0.4, 0.0];
blue_map = [linspace(light_blue(1), dark_blue(1), 128)', linspace(light_blue(2), dark_blue(2), 128)', linspace(light_blue(3), dark_blue(3), 128)'];
red_map = [linspace(light_red(1), dark_red(1), 128)', linspace(light_red(2), dark_red(2), 128)', linspace(light_red(3), dark_red(3), 128)'];
green_map = [linspace(light_green(1), dark_green(1), 128)', linspace(light_green(2), dark_green(2), 128)', linspace(light_green(3), dark_green(3), 128)'];
max_dom_g1=0; max_dom_g2=0; max_dom_g3=0;
for i=1:num_intervals; map_i=mechanic_counts_over_time{i}; if map_i.Count > 0; all_bay_counts=values(map_i); for k=1:length(all_bay_counts); counts=all_bay_counts{k}; group_counts = counts(2:4); [dominant_count, dominant_idx] = max(group_counts); if dominant_idx == 1; max_dom_g1=max(max_dom_g1, dominant_count); elseif dominant_idx == 2; max_dom_g2=max(max_dom_g2, dominant_count); elseif dominant_idx == 3; max_dom_g3=max(max_dom_g3, dominant_count); end; end; end; end
avg_max_g1=max(0.1, max_dom_g1/num_days); avg_max_g2=max(0.1, max_dom_g2/num_days); avg_max_g3=max(0.1, max_dom_g3/num_days);

% --- Create Figure, Main Axes, and Final Colorbar Layout ---
fig = figure('Name', 'Animated Mechanic Heatmap', 'Position', [100, 100, 1000, 800]);
main_ax = axes('Position', [0.05 0.2 0.9 0.75]); % Adjust main axes to make room at the bottom

% Create three horizontal colorbars arranged side-by-side
cb1_ax = axes('Position', [0.10, 0.08, 0.25, 0.03]); 
colormap(cb1_ax, blue_map); 
cb1 = colorbar(cb1_ax, 'Location','southoutside'); 
clim(cb1_ax, [0, avg_max_g1]); 
title('Local (G1)');

cb2_ax = axes('Position', [0.37, 0.08, 0.25, 0.03]); 
colormap(cb2_ax, red_map); 
cb2 = colorbar(cb2_ax, 'Location','southoutside'); 
clim(cb2_ax, [0, avg_max_g2]); 
title('CNAC (G2)');

cb3_ax = axes('Position', [0.64, 0.08, 0.25, 0.03]); 
colormap(cb3_ax, green_map); 
cb3 = colorbar(cb3_ax, 'Location','southoutside'); 
clim(cb3_ax, [0, avg_max_g3]); 
title('CI+Others (G3)');

axis([cb1_ax, cb2_ax, cb3_ax], 'off'); % Turn off the invisible axes boxes

% --- Create a lookup map for bay locations ---
bay_location_map = containers.Map('KeyType','char','ValueType','any');
all_bay_locs = {};
for r=1:size(raw_map,1); for c=1:size(raw_map,2); val=raw_map{r,c}; if ischar(val)||isstring(val); bay=upper(strtrim(string(val))); if endsWith(bay,"L")||endsWith(bay,"R"); bay=extractBefore(bay,strlength(bay)); end; if ~isempty(regexp(bay,'^[A-Z]+\d+$','once')); bay_location_map(bay)=[r,c]; all_bay_locs{end+1} = [r,c]; end; end; end; end

% === Frame Loop ===
for frame_idx = 1:num_intervals
    current_map = mechanic_counts_over_time{frame_idx};
    [rows, cols] = size(raw_map);
    rgb_map = ones(rows, cols, 3) * 0.8;
    numeric_map_totals = nan(rows, cols);
    for i = 1:length(all_bay_locs); loc = all_bay_locs{i}; rgb_map(loc(1), loc(2), :) = [1, 1, 1]; end
    active_bays = keys(current_map);
    for i = 1:length(active_bays)
        bay_name = active_bays{i};
        if isKey(bay_location_map, bay_name)
            loc = bay_location_map(bay_name); r = loc(1); c = loc(2);
            counts = current_map(bay_name);
            numeric_map_totals(r, c) = counts(1);
            group_counts = counts(2:4);
            [dominant_count, dominant_idx] = max(group_counts);
            if dominant_count > 0
                avg_dominant_count = dominant_count / num_days;
                if dominant_idx == 1; norm_val=min(1,avg_dominant_count/avg_max_g1); cmap=blue_map; elseif dominant_idx == 2; norm_val=min(1,avg_dominant_count/avg_max_g2); cmap=red_map; else; norm_val=min(1,avg_dominant_count/avg_max_g3); cmap=green_map; end
                color_idx = round(norm_val * (size(cmap,1)-1)) + 1;
                rgb_map(r, c, :) = cmap(max(1, color_idx),:);
            end
        end
    end
    axes(main_ax); cla;
    image(rgb_map);
    axis equal tight;
    set(gca, 'XTick', [], 'YTick', []);
    mins = (frame_idx - 1) * 5;
    title(sprintf("Daily Avg Mechanics by Dominant Group | Time %02d:%02d - %02d:%02d", floor(mins/60), mod(mins,60), floor((mins+5)/60), mod((mins+5),60)));
    total_all=0; total_g1=0; total_g2=0; total_g3=0;
    if current_map.Count>0; all_counts=values(current_map); for i=1:length(all_counts); c_arr=all_counts{i}; total_all=total_all+c_arr(1); total_g1=total_g1+c_arr(2); total_g2=total_g2+c_arr(3); total_g3=total_g3+c_arr(4); end; end
    avg_all=total_all/num_days; avg_g1=total_g1/num_days; avg_g2=total_g2/num_days; avg_g3=total_g3/num_days;
    stats_string = sprintf('Avg. Total: %0.1f (Local: %0.1f, CNAC: %0.1f, CI+Others: %0.1f)', avg_all, avg_g1, avg_g2, avg_g3);
    xlabel(stats_string, 'FontSize', 10, 'HorizontalAlignment', 'center');
    for r=1:rows; for c=1:cols; total_val=numeric_map_totals(r,c); if ~isnan(total_val) && total_val > 0; avg_val=total_val/num_days; text(c,r,sprintf('%0.1f',avg_val),'HorizontalAlignment','center','Color','black','FontSize',8,'FontWeight','bold');end;end;end
    drawnow;
    frame = getframe(fig);
    writeVideo(video_writer, frame);
end
close(video_writer);
close(fig);
disp("✔ Video 'mechanic_heatmap.mp4' created successfully!");

%% === Helper Functions ===
function [dt, minutes] = parse_time(entry_string)
    day_offset = 0;
    if contains(entry_string, '+') || contains(entry_string, '-'); parts = split(entry_string, {'+', '-'}); time_part = parts{1}; offset_part = parts{2}; sign_val = contains(entry_string, '+')*1 + contains(entry_string, '-')*(-1); day_offset = sign_val * str2double(offset_part); else; time_part = entry_string; end
    time_part = sprintf('%04s', time_part);
    dt = datetime(time_part, 'InputFormat', 'HHmm') + days(day_offset);
    minutes = hour(dt) * 60 + minute(dt);
end

%%% --- ADDED HELPER FUNCTION ---
function counts_over_time = apply_rule(counts_over_time, bay_name, report_group, base_minutes, time_offsets, num_mech, num_intervals)
    % Applies a single mechanic rule to the master count variable
    start_min = base_minutes + time_offsets(1);
    end_min = base_minutes + time_offsets(2);
    start_idx = floor(start_min / 5) + 1;
    end_idx = floor((end_min - 1) / 5) + 1;
    for t = start_idx:end_idx
        actual_idx = mod(t - 1, num_intervals) + 1;
        map_t = counts_over_time{actual_idx};
        if isKey(map_t, bay_name); current_counts = map_t(bay_name); else; current_counts = [0, 0, 0, 0]; end
        current_counts(1) = current_counts(1) + num_mech;
        if report_group == 1; current_counts(2) = current_counts(2) + num_mech; elseif report_group == 2; current_counts(3) = current_counts(3) + num_mech; elseif report_group == 3; current_counts(4) = current_counts(4) + num_mech; end
        map_t(bay_name) = current_counts;
        counts_over_time{actual_idx} = map_t;
    end
end

% %% --- Generate and Save a Comparison Chart of Mechanic Distribution ---
% disp('Generating planned vs. actual mechanic distribution chart...');
% 
% % =========================================================================
% % STEP 1: CALCULATE 'ACTUAL' MECHANIC DISTRIBUTION (NEW LOGIC)
% % This section remains unchanged as it already calculates the necessary
% % breakdown of mechanics by reporting group (Total, G1, G2, G3).
% % =========================================================================
% 
% % Create a new, separate variable to store the 'actual' counts
% actual_mechanic_counts = cell(num_intervals, 1);
% for i = 1:num_intervals
%     actual_mechanic_counts{i} = containers.Map('KeyType', 'char', 'ValueType', 'any');
% end
% 
% % Repeat the processing loop for both sheets
% for pass = 1:2
%     if pass == 1; sheet_name = arrival_sheet; is_departure = false; else; sheet_name = departure_sheet; is_departure = true; end
%     fprintf('Processing ACTUAL rules from sheet: %s...\n', sheet_name);
% 
%     opts = detectImportOptions(data_file, 'Sheet', sheet_name);
%     opts.DataRange = 'A2';
%     opts.SelectedVariableNames = [3, 9, 16, 17]; 
%     data = readtable(data_file, opts);
% 
%     for j = 1:height(data)
%         scheduled_entry = string(data{j,1});
%         actual_entry = string(data{j,2});
%         bay = string(data{j,3});
%         airline_code = upper(strtrim(string(data{j,4})));
% 
%         if ismissing(scheduled_entry) || ismissing(actual_entry) || ismissing(bay) || ismissing(airline_code)
%             continue;
%         end
%         if ~isKey(rule_group_map, airline_code) || ~isKey(report_group_map, airline_code)
%             continue;
%         end
% 
%         try
%             [~, scheduled_time_in_minutes] = parse_time(scheduled_entry);
%             [actual_time, actual_time_in_minutes] = parse_time(actual_entry);
%             scheduled_time = parse_time(scheduled_entry); % We need the datetime object for comparison
%         catch
%             continue;
%         end
% 
%         rule_group_num = rule_group_map(airline_code);
%         report_group_num = report_group_map(airline_code);
%         bay_name = upper(strtrim(bay));
%         if endsWith(bay_name, "L") || endsWith(bay_name, "R")
%             bay_name = extractBefore(bay_name, strlength(bay_name));
%         end
% 
%         times = {}; personnel = [];
%         if ~is_departure
%             switch rule_group_num; case 1; times = {[-20, 20], [20, 35]}; personnel = [2, 1]; case 2; times = {[-25, 35]}; personnel = [1]; case 3; times = {[-20, 35]}; personnel = [1]; case 4; times = {[-20, 35]}; personnel = [1]; case 5; times = {[-20, 35]}; personnel = [1]; case 6; times = {[-20, 25]}; personnel = [1]; end
%         else
%             switch rule_group_num; case 1; times = {[-70, -25], [-25, 15]}; personnel = [1, 2]; case 2; times = {[-70, -15], [-15, 15]}; personnel = [1, 2]; case 3; times = {[-70, -20], [-20, 15]}; personnel = [1, 2]; case 4; times = {[-70, -15], [-15, 15]}; personnel = [1, 2]; case 5; times = {[-70, -15], [-15, 15]}; personnel = [1, 2]; case 6; times = {[-70, -20], [-20, 15]}; personnel = [1, 2]; end
%         end
% 
%         if actual_time > scheduled_time
%             time_diff_minutes = minutes(actual_time - scheduled_time);
%             delay_extension = 5 * ceil(time_diff_minutes / 5);
%             if rule_group_num == 1 && ~is_departure
%                 actual_mechanic_counts = apply_rule(actual_mechanic_counts, bay_name, report_group_num, scheduled_time_in_minutes, times{1}, personnel(1), num_intervals);
%                 ext_start_min = scheduled_time_in_minutes + times{1}(2);
%                 ext_end_min = ext_start_min + delay_extension;
%                 actual_mechanic_counts = apply_rule(actual_mechanic_counts, bay_name, report_group_num, 0, [ext_start_min, ext_end_min], personnel(1), num_intervals);
%                 shift = times{2}(2) - times{2}(1);
%                 shifted_start_min = ext_end_min;
%                 shifted_end_min = shifted_start_min + shift;
%                 actual_mechanic_counts = apply_rule(actual_mechanic_counts, bay_name, report_group_num, 0, [shifted_start_min, shifted_end_min], personnel(2), num_intervals);
%             else
%                 for k = 1:length(times)
%                     actual_mechanic_counts = apply_rule(actual_mechanic_counts, bay_name, report_group_num, scheduled_time_in_minutes, times{k}, personnel(k), num_intervals);
%                 end
%                 last_rule_end_min = scheduled_time_in_minutes + times{end}(2);
%                 ext_start_min = last_rule_end_min;
%                 ext_end_min = ext_start_min + delay_extension;
%                 actual_mechanic_counts = apply_rule(actual_mechanic_counts, bay_name, report_group_num, 0, [ext_start_min, ext_end_min], personnel(end), num_intervals);
%             end
%         else
%             rounded_down_minutes = 5 * floor(actual_time_in_minutes / 5);
%             rounded_up_minutes   = 5 * ceil(actual_time_in_minutes / 5);
%             for k = 1:length(times)
%                 start_min = rounded_down_minutes + times{k}(1);
%                 end_min = rounded_up_minutes + times{k}(2);
%                 actual_mechanic_counts = apply_rule(actual_mechanic_counts, bay_name, report_group_num, 0, [start_min, end_min], personnel(k), num_intervals);
%             end
%         end
%     end
% end
% 
% % =========================================================================
% % STEP 2: CALCULATE AVERAGE DISTRIBUTIONS FOR PLOTTING
% % =========================================================================
% 
% % --- Calculate 'PLANNED' average distribution (from main script variable) ---
% avg_planned_mechanics = zeros(1, num_intervals);
% for t = 1:num_intervals
%     map_t = mechanic_counts_over_time{t}; % Uses the original 'scheduled' data
%     total = 0;
%     if map_t.Count > 0
%         all_counts = values(map_t);
%         for k = 1:length(all_counts); total = total + all_counts{k}(1); end
%     end
%     avg_planned_mechanics(t) = total / num_days;
% end
% 
% % --- MODIFIED: Calculate 'ACTUAL' average distribution for EACH reporting group ---
% avg_actual_g1 = zeros(1, num_intervals);
% avg_actual_g2 = zeros(1, num_intervals);
% avg_actual_g3 = zeros(1, num_intervals);
% 
% for t = 1:num_intervals
%     map_t = actual_mechanic_counts{t}; % Uses the new 'actual' data
%     total_g1 = 0; total_g2 = 0; total_g3 = 0;
%     if map_t.Count > 0
%         all_counts = values(map_t);
%         for k = 1:length(all_counts)
%             counts_array = all_counts{k};
%             total_g1 = total_g1 + counts_array(2);
%             total_g2 = total_g2 + counts_array(3);
%             total_g3 = total_g3 + counts_array(4);
%         end
%     end
%     avg_actual_g1(t) = total_g1 / num_days;
%     avg_actual_g2(t) = total_g2 / num_days;
%     avg_actual_g3(t) = total_g3 / num_days;
% end
% 
% 
% % =========================================================================
% % STEP 3: CREATE AND FORMAT THE COMPARISON CHART (REWRITTEN)
% % =========================================================================
% figure('Name', 'Mechanic Distribution Comparison');
% hold on;
% 
% % --- Create the data matrix for the stacked area chart ---
% % Each column is a category, each row is a time interval
% stacked_data = [avg_actual_g1', avg_actual_g2', avg_actual_g3'];
% 
% % --- Plot the stacked area chart for the 'Actual' distribution ---
% actual_stacked_area = area(stacked_data);
% 
% % --- Overlay the 'Planned' distribution as a single black line ---
% planned_line = plot(avg_planned_mechanics, 'Color', 'k', 'LineWidth', 2);
% 
% hold off;
% 
% % --- Formatting and Customization ---
% grid on;
% title('Planned vs. Actual Daily Average Mechanic Distribution', 'FontSize', 14);
% ylabel('Average Number of Mechanics', 'FontSize', 12);
% xlabel('Time of Day', 'FontSize', 12);
% xlim([1, num_intervals]);
% 
% % Customize the colors of the stacked areas
% actual_stacked_area(1).FaceColor = [0.2, 0.6, 1.0]; % Blue for Local
% actual_stacked_area(2).FaceColor = [1.0, 0.4, 0.4]; % Red for CNAC
% actual_stacked_area(3).FaceColor = [0.5, 0.8, 0.5]; % Green for CI + Others
% 
% % Add a legend to identify all chart components
% legend([planned_line, actual_stacked_area], ...
%     {'Planned (Total)', 'Actual (Local Mechanics)', 'Actual (CNAC)', 'Actual (CI + Others)'}, ...
%     'Location', 'northwest');
% 
% % --- Create readable time labels for the X-axis ---
% tick_positions = 1:24:num_intervals;
% tick_labels = {};
% for i = 0:2:23; tick_labels{end+1} = sprintf('%02d:00', i); end
% xticks(tick_positions);
% xticklabels(tick_labels);
% 
% % --- Save the figure ---
% try
%     saveas(gcf, 'mechanics_comparison_chart.png');
%     fprintf("✔ Chart saved successfully as 'mechanics_comparison_chart.png'\n");
% catch ME
%     fprintf("Could not save the chart. Error: %s\n", ME.message);
% end
% 
% %% ========================================================================
% % HELPER FUNCTIONS (Included for completeness)
% % =========================================================================
% function [dt, minutes] = parse_time(entry_string)
%     day_offset = 0;
%     if contains(entry_string, '+') || contains(entry_string, '-'); parts = split(entry_string, {'+', '-'}); time_part = parts{1}; offset_part = parts{2}; sign_val = contains(entry_string, '+')*1 + contains(entry_string, '-')*(-1); day_offset = sign_val * str2double(offset_part); else; time_part = entry_string; end
%     time_part = sprintf('%04s', time_part);
%     dt = datetime(time_part, 'InputFormat', 'HHmm') + days(day_offset);
%     minutes = hour(dt) * 60 + minute(dt);
% end
% function counts_over_time = apply_rule(counts_over_time, bay_name, report_group, base_minutes, time_offsets, num_mech, num_intervals)
%     start_min = base_minutes + time_offsets(1);
%     end_min = base_minutes + time_offsets(2);
%     start_idx = floor(start_min / 5) + 1;
%     end_idx = floor((end_min - 1) / 5) + 1;
%     for t = start_idx:end_idx
%         actual_idx = mod(t - 1, num_intervals) + 1;
%         map_t = counts_over_time{actual_idx};
%         if isKey(map_t, bay_name); current_counts = map_t(bay_name); else; current_counts = [0, 0, 0, 0]; end
%         current_counts(1) = current_counts(1) + num_mech;
%         if report_group == 1; current_counts(2) = current_counts(2) + num_mech; elseif report_group == 2; current_counts(3) = current_counts(3) + num_mech; elseif report_group == 3; current_counts(4) = current_counts(4) + num_mech; end
%         map_t(bay_name) = current_counts;
%         counts_over_time{actual_idx} = map_t;
%     end
% end

% %%
% % --- Find and Display Peak Airport-Wide Mechanic Interval (ACTUAL) ---
% disp('Calculating ACTUAL peak mechanic statistics...');
% 
% % Initialize variables to track the peak
% peak_total_mechanics = 0;
% peak_time_intervals = [];
% 
% % Loop through each 5-minute interval of the day
% for t = 1:num_intervals
%     % *** MODIFIED LINE: Use the 'actual' data variable ***
%     map_t = actual_mechanic_counts{t};
% 
%     current_interval_total = 0;
% 
%     % If there are any mechanics at this interval, sum them all up
%     if map_t.Count > 0
%         all_bay_counts_in_map = values(map_t);
%         for k = 1:length(all_bay_counts_in_map)
%             % The value is a 1x4 array; the first element is the total
%             counts_array = all_bay_counts_in_map{k};
%             current_interval_total = current_interval_total + counts_array(1);
%         end
%     end
% 
%     % Check if this interval's total is a new peak
%     if current_interval_total > peak_total_mechanics
%         peak_total_mechanics = current_interval_total;
%         peak_time_intervals = t; % New peak found, so reset the list of times
%     elseif current_interval_total == peak_total_mechanics && peak_total_mechanics > 0
%         peak_time_intervals = [peak_time_intervals, t]; % A tie was found, so add it to the list
%     end
% end
% 
% % Calculate the daily average for the peak time
% avg_peak_mechanics = peak_total_mechanics / num_days;
% 
% % Display the final results
% fprintf('\n--- Peak ACTUAL Mechanic Statistics (Airport-Wide) ---\n');
% if ~isempty(peak_time_intervals)
%     fprintf('Busiest 5-Minute Interval(s) (Actual):\n');
%     for t = peak_time_intervals
%         mins = (t - 1) * 5;
%         time_str = sprintf('%02d:%02d - %02d:%02d', ...
%             floor(mins/60), mod(mins,60), floor((mins+5)/60), mod((mins+5),60));
%         fprintf('  - %s\n', time_str);
%     end
%     fprintf('Total Actual Mechanics during peak interval(s): %d\n', peak_total_mechanics);
%     fprintf('Daily Average Actual Mechanics during peak interval(s): %0.1f\n', avg_peak_mechanics);
% else
%     fprintf('No mechanic activity was recorded in the actual data.\n');
% end

