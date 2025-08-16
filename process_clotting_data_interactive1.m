function process_clotting_data_interactive1()
close all
    try
        % Let the user select the Excel file
        [file_name, path_name] = uigetfile('*.xlsx', 'Select the Excel file');
        if isequal(file_name, 0)
            disp('User selected Cancel');
            return;
        else
            file_path = fullfile(path_name, file_name);
        end
        
        % Get the available sheet names
        sheets = sheetnames(file_path);
        
        % Prompt the user to enter the sheet index
        [sheet_index, ok] = listdlg('PromptString', 'Select a sheet:', ...
                                    'SelectionMode', 'single', ...
                                    'ListString', sheets);
        if ~ok
            disp('User cancelled sheet selection');
            return;
        end
        
        % Read the table from the specified Excel sheet
        table = readtable(file_path, 'Sheet', sheets{sheet_index}, 'VariableNamingRule', 'preserve');
        

        % Store headers separately
        headers = table.Properties.VariableNames;
        
        % Remove the first row (containing numbers for easier identification)
        table(1, :) = [];
        
        % Convert table to array for data analysis
        data = table2array(table);
        t = data(:, 1);
        
        % Extract base names of the selected groups without duplicates
        base_names = cellfun(@(x) strtok(x, '_'), headers(2:end), 'UniformOutput', false);
        unique_groups = unique(base_names);
        
        % Prompt the user to select which groups to use
        [selected_groups_idx, ok] = listdlg('PromptString', 'Select groups to include:', ...
                                            'ListString', unique_groups, ...
                                            'SelectionMode', 'multiple');
        if ~ok
            disp('User cancelled group selection');
            return;
        end
        
        x_label_input = sheets{sheet_index};

        % Get the selected group names
        selected_groups = unique_groups(selected_groups_idx);
        
        % Find columns that belong to the selected groups
        selected_cols_idx = find(ismember(base_names, selected_groups));
        
        % Filter the headers and data based on the user selection
        selected_headers = headers(selected_cols_idx + 1); % Adjust for time column
        selected_data = data(:, selected_cols_idx + 1); % Adjust for time column
        
        % Group the selected columns based on their base names
        [~, ~, idx] = unique(cellfun(@(x) strtok(x, '_'), selected_headers, 'UniformOutput', false));
        
        % Initialize group columns
        group_cols = cell(numel(selected_groups), 1);
        
        % Assign columns to groups based on the selected headers
        for i = 1:numel(selected_groups)
            group_cols{i} = find(idx == i);
        end
        
        % Calculate means and stds for each group
        groups = cell(numel(selected_groups), 1);
        means = zeros(length(t), numel(selected_groups));
        stds = zeros(length(t), numel(selected_groups));
        for i = 1:numel(selected_groups)
            group_data = selected_data(:, group_cols{i});
            means(:, i) = mean(group_data, 2);
            stds(:, i) = std(group_data, 0, 2);
            groups{i} = group_data;
        end

        % Determine number of groups for color assignment
        num_groups = numel(selected_groups);
        cmap = lines(num_groups);  % Base colormap for fallback colors

        % Define a color map with distinct, bright, and color-blind friendly colors
        group_colors = containers.Map();

        % Group 1: Bare, Heparin, Heprasil, Thiol, Negative
        group_colors('bare') = [173, 216, 230] / 255;  % Lighter Bright Blue
        group_colors('heparin') = [255, 153, 153] / 255;  % Lighter Bright Red
        group_colors('heprasil') = [144, 238, 144] / 255;  % Lighter Bright Green
        group_colors('thiol') = [224, 176, 255] / 255;  % Lighter Magenta
        group_colors('negative') = [255, 204, 153] / 255;  % Lighter Orange

        % Group 2: Bare, Negative, Eluting, Collagen
        group_colors('eluting') = [204, 153, 255] / 255;  % Lighter Purple
        group_colors('collagen') = [255, 255, 153] / 255;  % Lighter Bright Yellow

        % Group 3: Bare, Eluting, Positive, Negative, Graft, Hybrid
        group_colors('positive') = [144, 238, 144] / 255;  % Lighter Lime Green
        group_colors('graft') = [204, 255, 255] / 255;  % Lighter Cyan
        group_colors('hybrid') = [255, 153, 153] / 255;  % Lighter Dark Red

         % Group 4: PPP Citrate Hybrid, PPP Heparin Hybrid, PPP EDTA Hybrid, PRP Citrate Hybrid
        group_colors('PPP EDTA Hybrid') = [173, 216, 230] / 255;  % Lighter Bright Blue
        group_colors('PPP Citrate Hybrid') = [255, 153, 153] / 255;  % Lighter Bright Red
        group_colors('PPP Heparin Hybrid') = [144, 238, 144] / 255;  % Lighter Bright Green
        group_colors('PRP Citrate Hybrid') = [224, 176, 255] / 255;  % Lighter Magenta
       
        % Assign dark colors for mean lines
        group_dark_colors = containers.Map();
        group_dark_colors('bare') = [55, 126, 184] / 255;  % Bright Blue
        group_dark_colors('heparin') = [228, 26, 28] / 255;  % Bright Red
        group_dark_colors('heprasil') = [77, 175, 74] / 255;  % Bright Green
        group_dark_colors('thiol') = [154, 1, 205] / 255;  % Dark Magenta
        group_dark_colors('negative') = [204, 100, 0] / 255;  % Dark Orange
        group_dark_colors('eluting') = [148, 0, 211] / 255;  % Dark Purple
        group_dark_colors('collagen') = [204, 204, 0] / 255;  % Dark Yellow
        group_dark_colors('positive') = [50, 205, 50] / 255;  % Lime Green
        group_dark_colors('graft') = [0, 183, 235] / 255;  % Cyan
        group_dark_colors('hybrid') = [139, 0, 0] / 255;  % Dark Red
        group_dark_colors('PPP EDTA Hybrid') = [55, 126, 184] / 255;  % Bright Blue
        group_dark_colors('PPP Citrate Hybrid') = [228, 26, 28] / 255;  % Bright Red
        group_dark_colors('PPP Heparin Hybrid') = [77, 175, 74] / 255;  % Bright Green
        group_dark_colors('PRP Citrate Hybrid') = [154, 1, 205] / 255;  % Dark Magenta

        % Assign colors to selected groups
        colors = cell(num_groups, 1);
        dark_colors = cell(num_groups, 1);

        for i = 1:num_groups
            group_name = selected_groups{i};
            if isKey(group_colors, group_name)
                colors{i} = group_colors(group_name);
                dark_colors{i} = group_dark_colors(group_name);
            else
                % Use distinct colors from the colormap for unknown groups
                dark_colors{i} = cmap(i, :);
                colors{i} = 0.5 * cmap(i, :) + 0.5;  % Lighter fill color
            end
        end

        % Determine subplot layout based on number of selected groups
        num_cols = ceil(sqrt(num_groups));
        num_rows = ceil(num_groups / num_cols);
        
        % Plotting
        figure(1)
        set(gcf, 'Name', ['Figure 1 (' sheets{sheet_index} ')'], 'NumberTitle', 'off', ...
            'Position', [100, 100, 1200, 800]);
for i = 1:num_groups
    subplot(num_rows, num_cols, i)
    hold on
    
    % Calculate x-axis limits based on the first NaN or empty value
    nan_index = find(any(isnan(data), 2), 1);
    if isempty(nan_index)
        x_max = max(t); % No NaN found, use the full range of t
    else
        x_max = t(max(1, nan_index-1)); % Use the time value before the first NaN
    end
    x_max = max(x_max, min(t)); % Ensure x_max is valid
    
    % Slice the time (`t`), means, and stds for the valid range
    valid_indices = t >= min(t) & t <= x_max;
    t_valid = t(valid_indices);
    means_valid = means(valid_indices, i);
    stds_valid = stds(valid_indices, i);
    
    % Create the filled area for standard deviation
    fill([t_valid; flipud(t_valid)], ...
         [means_valid + stds_valid; flipud(means_valid - stds_valid)], ...
         colors{i}, 'FaceAlpha', 0.5, 'EdgeColor', 'none');
    
    % Plot the mean line
    plot(t_valid, means_valid, '-', 'color', dark_colors{i}, 'LineWidth', 1.5)
    
    % Label axes and set the title
    xlabel('Time (m)','FontSize',17)
    ylabel('Absorbance','FontSize',17)
    title(selected_groups{i},'FontSize',14)
    
    % Set axis limits
    xlim([min(t), x_max]);
    ylim([0, (max(means_valid)+max(stds_valid)+0.15)]);
    
    % Add grid and box
    box on
    grid on
    hold off
end

%% 2

        % Combined plot
        figure(2)
        set(gcf, 'Position', [100, 100, 1200, 1000]);
        h1= subplot(2, 1, 1);
        set(h1, 'Position', [0.1, 0.65, 0.8, 0.30]);
        hold on
        for i = 1:num_groups
            plot(t, means(:, i), '-', 'Color', dark_colors{i}, 'LineWidth', 1.5)
        end
        legend(selected_groups, 'Location', 'best', 'FontSize', 14)
        xlabel('Time (m)')
        ylabel('Absorbance')
        title(sheets{sheet_index}, 'FontSize',20)
        % Determine the maximum x-limit based on the first NaN or empty value in data
            xlim([min(t), x_max]);
        ylim([0, 1.35]);
        set(gca, 'Fontsize', 25, 'Linewidth', 1.5)
        box on
        grid on
        hold off

% Check mean data line for clotting time before individual measurements
mean_clotting_times = zeros(1, num_groups);

% Identify if "hybrid" is the only group that clots
is_hybrid_only = false;
non_clotting_groups = true(1, num_groups); % Assume all groups do not clot initially

for i = 1:num_groups
    % Find the time where the mean exceeds the threshold (10% above initial)
    threshold = means(1, i) * 1.10; % 10% above the initial mean value
    I = find(means(:, i) > threshold, 1); % First point exceeding the threshold
    if isempty(I)
        mean_clotting_times(i) = max(t); % No clotting detected
    else
        mean_clotting_times(i) = t(I); % Time to clot based on mean
        non_clotting_groups(i) = false; % Mark group as clotting
    end
end

% Check if "hybrid" is the only group that clots
if sum(~non_clotting_groups) == 1
    hybrid_idx = find(strcmp(selected_groups, 'hybrid'));
    if ~isempty(hybrid_idx) && ~non_clotting_groups(hybrid_idx)
        is_hybrid_only = true;
    end
end

% If "hybrid" is the only group that clots, treat it as non-clotting
if is_hybrid_only
    mean_clotting_times(:) = max(t); % Set all times to max(t)
end

% Display mean clotting times
disp('Mean clotting times (based on group averages):');
disp(mean_clotting_times);


% Calculate individual clotting times
clotting_times = struct();
for i = 1:num_groups
    X = selected_data(:, group_cols{i});
    n = size(X, 2);
    times = zeros(n, 1);
    for j = 1:n
        column = X(:, j);
        I = find(column > column(1) * 1.10);
        if isempty(I)
            times(j) = max(t); % Default to max time if no clotting
        else
            times(j) = data(I(1), 1); % Clotting time based on threshold
        end
    end
    clotting_times.(selected_groups{i}) = times;
end

% If no valid mean clotting times are detected, use max(t) for plotting
if all(mean_clotting_times >= max(t))
    for i = 1:num_groups
        clotting_times.(selected_groups{i}) = max(t) * ones(size(clotting_times.(selected_groups{i})));
    end
end

% Plot individual clotting times
figure(2)
%set(h1, 'Position', [0.1, 0.6, 0.8, 0.35]); % Top subplot: move up slightly and reduce height
%set(h2, 'Position', [0.1, 0.1, 0.8, 0.5]); % Bottom subplot: increase height

h2= subplot(2, 1, 2);
hold on
bar_heights = zeros(1, num_groups);
for i = 1:num_groups
    times = clotting_times.(selected_groups{i});
    bar_heights(i) = mean(times);
    scatter(i * ones(size(times)), times, 'linewidth', 2, 'MarkerEdgeColor', dark_colors{i});
end

set(h2, 'Position', [0.1, 0.1, 0.8, 0.45]); % Bottom subplot: increase height

% Create the bar chart
for i = 1:numel(bar_heights)
    bar(i, bar_heights(i), 'EdgeColor', dark_colors{i}, 'FaceColor', 'none', 'LineWidth', 2);
end

% Add custom x-axis labels
xticks(1:num_groups)
xticklabels(selected_groups)
xtickangle(45)
ylim([0, max(cellfun(@max, struct2cell(clotting_times))) + 20]); % Add margin above bars for annotations


% Prepare data for ANOVA
y_anova = [];
group_labels = [];
for i = 1:num_groups
    times = clotting_times.(selected_groups{i});
    y_anova = [y_anova; times];
    group_labels = [group_labels; repmat(selected_groups(i), length(times), 1)];
end

% Ensure group_labels is a categorical variable
group_labels = categorical(group_labels);

% Define control groups
positive_control = 'positive';
negative_control = 'negative';

% Get indices of controls
positive_idx = find(strcmp(selected_groups, positive_control));
negative_idx = find(strcmp(selected_groups, negative_control));

% Perform ANOVA and post-hoc analysis if enough data
if numel(unique(group_labels)) > 1
    [p, ~, stats] = anova1(y_anova, group_labels, 'off'); % Run ANOVA
    [c, ~, ~, ~] = multcompare(stats, 'Display', 'off'); % Multiple comparisons
else
    disp('Not enough unique groups to perform ANOVA.');
    c = []; % Define `c` as empty to avoid errors later
end

% Symbols for significance
group_symbols = containers.Map();
group_symbols(positive_control) = 'x';
group_symbols(negative_control) = 'o';

% Assign unique symbols to other groups
all_groups = selected_groups(~ismember(selected_groups, {positive_control, negative_control}));
unique_symbols = ['^', '#', 's', 'd', '*']; % Symbols for other groups
for i = 1:numel(all_groups)
    group_symbols(all_groups{i}) = unique_symbols(mod(i-1, numel(unique_symbols)) + 1);
end

% Display group-symbol mapping in the console
disp('Group-symbol mapping:');
disp(group_symbols);

% Calculate the maximum height across all groups
max_total_height = max(cellfun(@max, struct2cell(clotting_times)));
y_position_base = max_total_height + 10; % Base y-position for annotations
y_step = 5; % Increment for each additional level of significance

if ~isempty(c)
    hold on;
    for row = 1:size(c, 1)
        group1 = c(row, 1); % First group index
        group2 = c(row, 2); % Second group index
        p_value = c(row, 6); % P-value for the comparison

        % Determine the level of significance
        if p_value < 0.001
            significance_level = 3; % Highly significant
        elseif p_value < 0.01
            significance_level = 2; % Moderately significant
        elseif p_value < 0.05
            significance_level = 1; % Weakly significant
        else
            continue; % Skip non-significant results
        end

        % Determine the target group and symbol
        if group1 == positive_idx || group2 == positive_idx
            % Comparison with positive control
            if group1 == positive_idx
                target_group = group2;
            else
                target_group = group1;
            end
            annotation_symbol = repmat(group_symbols(positive_control), 1, significance_level);
            y_position = y_position_base + y_step; % Use higher position for positive control
        elseif group1 == negative_idx || group2 == negative_idx
            % Comparison with negative control
            if group1 == negative_idx
                target_group = group2;
            else
                target_group = group1;
            end
            annotation_symbol = repmat(group_symbols(negative_control), 1, significance_level);
            y_position = y_position_base; % Use base position for negative control
        else
            % Comparison outside of control groups
            disp(['Significant comparison between ', selected_groups{group1}, ...
                  ' and ', selected_groups{group2}, ' with p-value: ', num2str(p_value)]);
            if group1 < group2
                target_group = group2; % Assign group2 as the target
            else
                target_group = group1; % Assign group1 as the target
            end
            annotation_symbol = repmat(group_symbols(selected_groups{group1}), 1, significance_level);
            y_position = y_position_base - y_step; % Adjust y-position
        end

        % Add annotation at the specified position with larger symbols
        text(target_group, y_position, annotation_symbol, ...
            'HorizontalAlignment', 'center', 'FontSize', 20, 'FontWeight', 'bold'); % Increased FontSize
    end
    hold off;
end

if isempty(c) || all(isnan(p)) % Ensure p is scalar or reduce it with `all`
    disp('No significant comparisons found or p-value is NaN.');

    % Add "ns" text above the plot
    text(mean(1:num_groups), y_position_base, 'ns', 'HorizontalAlignment', 'center', ...
        'FontSize', 14, 'FontWeight', 'bold', 'Color', 'k');
end



%% 
% Use the input as the x-label
xlabel(sheets{sheet_index});
ylabel('Time (m)');
%title('Time to Initiate Clotting');
set(gca, 'Fontsize', 20, 'Linewidth', 1.5)
hold off


   %% 3     
       figure(3)
hold on

if length(unique(group_labels)) > 1
    [p, tbl, stats] = anova1(y_anova, group_labels, 'off');
    % Display the p-value from ANOVA results
    disp(['p-value (ANOVA): ', num2str(p)])
    boxplot_handle = boxplot(y_anova, group_labels);

    % Customize box plot thickness
    set(findobj(boxplot_handle, 'Type', 'Line'), 'LineWidth', 1.5); % General box plot lines
    median_lines = findobj(boxplot_handle, 'Tag', 'Median'); % Median lines
    set(median_lines, 'LineWidth', 2.5, 'Color', 'r'); % Thicker red median lines

    xlabel(x_label_input);
    ylabel('Time (m)');
    ylim([0, max(cellfun(@max, struct2cell(clotting_times))) + 20]); % Add margin above bars for annotations
    title([ x_label_input]);
    set(gca, 'FontSize', 25, 'LineWidth', 1.5);

    % Perform multiple comparisons
    [c, ~, ~, ~] = multcompare(stats, 'Display', 'off');


% Perform ANOVA and post-hoc analysis if enough data
if numel(unique(group_labels)) > 1
    [p, ~, stats] = anova1(y_anova, group_labels, 'off'); % Run ANOVA
    [c, ~, ~, ~] = multcompare(stats, 'Display', 'off'); % Multiple comparisons
else
    disp('Not enough unique groups to perform ANOVA.');
    c = []; % Define `c` as empty to avoid errors later
end

% Symbols for significance
group_symbols = containers.Map();
group_symbols(positive_control) = 'x';
group_symbols(negative_control) = 'o';

% Assign unique symbols to other groups
all_groups = selected_groups(~ismember(selected_groups, {positive_control, negative_control}));
unique_symbols = ['^', '#', 's', 'd', '*']; % Symbols for other groups
for i = 1:numel(all_groups)
    group_symbols(all_groups{i}) = unique_symbols(mod(i-1, numel(unique_symbols)) + 1);
end

% Display group-symbol mapping in the console
disp('Group-symbol mapping:');
disp(group_symbols);

% Calculate the maximum height across all groups
max_total_height = max(cellfun(@max, struct2cell(clotting_times)));
y_position_base = max_total_height + 10; % Base y-position for annotations
y_step = 5; % Increment for each additional level of significance

if ~isempty(c)
    hold on;
    for row = 1:size(c, 1)
        group1 = c(row, 1); % First group index
        group2 = c(row, 2); % Second group index
        p_value = c(row, 6); % P-value for the comparison

        % Determine the level of significance
        if p_value < 0.001
            significance_level = 3; % Highly significant
        elseif p_value < 0.01
            significance_level = 2; % Moderately significant
        elseif p_value < 0.05
            significance_level = 1; % Weakly significant
        else
            continue; % Skip non-significant results
        end

        % Determine the target group and symbol
        if group1 == positive_idx || group2 == positive_idx
            % Comparison with positive control
            if group1 == positive_idx
                target_group = group2;
            else
                target_group = group1;
            end
            annotation_symbol = repmat(group_symbols(positive_control), 1, significance_level);
            y_position = y_position_base + y_step; % Use higher position for positive control
        elseif group1 == negative_idx || group2 == negative_idx
            % Comparison with negative control
            if group1 == negative_idx
                target_group = group2;
            else
                target_group = group1;
            end
            annotation_symbol = repmat(group_symbols(negative_control), 1, significance_level);
            y_position = y_position_base; % Use base position for negative control
        else
            % Comparison outside of control groups
            disp(['Significant comparison between ', selected_groups{group1}, ...
                  ' and ', selected_groups{group2}, ' with p-value: ', num2str(p_value)]);
            if group1 < group2
                target_group = group2; % Assign group2 as the target
            else
                target_group = group1; % Assign group1 as the target
            end
            annotation_symbol = repmat(group_symbols(selected_groups{group1}), 1, significance_level);
            y_position = y_position_base - y_step; % Adjust y-position
        end

        % Add annotation at the specified position with larger symbols
        text(target_group, y_position, annotation_symbol, ...
            'HorizontalAlignment', 'center', 'FontSize', 20, 'FontWeight', 'bold'); % Increased FontSize
    end
    hold off;
end

if isempty(c) || all(isnan(p)) % Ensure p is scalar or reduce it with `all`
    disp('No significant comparisons found or p-value is NaN.');

    % Add "ns" text above the plot
    text(mean(1:num_groups), y_position_base, 'NS', 'HorizontalAlignment', 'center', ...
        'FontSize', 20, 'FontWeight', 'bold', 'Color', 'k');
end


%% 4
            
            [c, m, h, nms] = multcompare(stats, 'Display', 'off');
figure(4)
hold on
axis off
title(['ANOVA Result with ', x_label_input]);

if p < 0.05
    text(0.5, 0.9, 'Significant Differences Detected', 'FontSize', 14, 'Color', 'g', 'HorizontalAlignment', 'center');
    text(0.5, 0.95, ['p-value (ANOVA): ', num2str(p)], 'FontSize', 14, 'Color', 'k', 'HorizontalAlignment', 'center');
else
    text(0.5, 0.9, 'No Significant Differences Detected', 'FontSize', 14, 'Color', 'r', 'HorizontalAlignment', 'center');
    text(0.5, 0.85, ['p-value (ANOVA): ', num2str(p)], 'FontSize', 12, 'Color', 'k', 'HorizontalAlignment', 'center');
end

% Extract significant and non-significant comparisons
significant_diff = c(c(:, 6) < 0.05, :);
not_significant_diff = c(c(:, 6) >= 0.05, :);

if isempty(significant_diff)
    text(0.5, 0.7, 'No significant differences found in post-hoc test.', 'FontSize', 14, 'Color', 'k', 'HorizontalAlignment', 'center');
else
    % Display significant differences
    for i = 1:size(significant_diff, 1)
        text(0.5, 0.7 - i * 0.05, ...
             ['Group ', nms{significant_diff(i, 1)}, ' is significantly different from Group ', nms{significant_diff(i, 2)}, ' with p-value ', num2str(significant_diff(i, 6))], ...
             'FontSize', 12, 'HorizontalAlignment', 'center');
    end
end

% Display non-significant comparisons
if isempty(not_significant_diff)
    text(0.5, 0.5 - size(significant_diff, 1) * 0.05 - 0.05, ...
         'All groups are significantly different.', 'FontSize', 14, 'Color', 'k', 'HorizontalAlignment', 'center');
else
    for i = 1:size(not_significant_diff, 1)
        text(0.5, 0.5 - (size(significant_diff, 1) + i) * 0.05, ...
             ['Group ', nms{not_significant_diff(i, 1)}, ' is not significantly different from Group ', nms{not_significant_diff(i, 2)}, ' with p-value ', num2str(not_significant_diff(i, 6))], ...
             'FontSize', 12, 'HorizontalAlignment', 'center');
    end
end
        end
hold off;

    catch ME
        disp('An error occurred:')
        disp(ME.message)
    end

    %% 
  % Specify the folder where you want to save the figures
output_folder = uigetdir(pwd, 'Select Folder to Save Figures');

% Check if the user selected a folder
if output_folder == 0
    disp('No folder selected. Figures will not be saved.');
else
    % Get all open figure handles
    fig_handles = findall(0, 'Type', 'figure');

    % Loop through each figure and save as PNG
    for i = 1:length(fig_handles)
        % Construct the filename
        fig_name = fullfile(output_folder, ['Figure' num2str(fig_handles(i).Number) '.png']);

        % Save the figure
        saveas(fig_handles(i), fig_name);
        disp(['Saved: ', fig_name]);
    end

    disp('All figures have been saved as PNGs.');

end
