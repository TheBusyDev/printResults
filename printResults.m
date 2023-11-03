function printResults (fileName, val)
    % syntax: printResults (fileName, val)
    %
    % Print text on both .txt, .xlsx files and on command prompt.
    % In .xlsx files, data are stored in the first sheet. Other sheets are not changed.
    %
    % 'fileName' must contain the name of the file where you want to store data.
    %
    % 'val' must be a 2-column matrix:
    % the first column contains text, 
    % the second column contains numerical values.
    % Different groups of data must be separated by "#".
    %
    % Example:
    %   val = [
    %       "#" "#"
    %       "Diameter" 50
    %       "Mass" 35
    %       "#" "#"
    %       "Diameter" 70
    %       "Mass" 90
    %       "#" "#"
    %       ...
    %   ];

    [row, col] = size (val);

    if col ~= 2
        error ('Input must be a 2-column matrix!')
    end

    xlsxName = sprintf ('%s.xlsx', fileName);
    txtName = sprintf('%s.txt', fileName);

    %% trim starting and ending '#'
    if val (1, 1) == "#"
        val = val (2:row, :);
        row = row-1;
    end

    if val (row, 1) == "#"
        row = row-1;
        val = val (1:row, :);
    end

    %% get number of labels
    len=0;

    while len < row && val (len+1, 1) ~= "#"
        len = len+1;
    end

    %% write labels on .xlsx file
    label = val (1:len, 1)'; % save labels of each data group
    writematrix (label, xlsxName, 'Sheet', 1, 'WriteMode', 'overwritesheet', 'AutoFitWidth', false);

    %% format data from 'val' and write them on .xlsx file
    i=2; % data are stored starting from the second row (the first row was previously filled with labels)
    j=1; % data are stored starting from the first column

    for k = 1:row
        if val (k, 1) ~= "#"
            cell = str2double (val (k, 2));

            if isinf (cell)
                if cell > 0
                    cell = "Inf";
                    val (k, 2) = "Inf";
                else
                    cell = "-Inf";
                    val (k, 2) = "-Inf";
                end
            elseif ~isnan (cell) % check if 'cell' is a floating-point number
                abs_num = abs (cell);
                
                if abs_num >= 10^4 || abs_num < 0.01 && abs_num ~= 0
                    val (k, 2) = sprintf ('%.3e', cell);
                elseif abs_num >= 10^3 || mod (abs_num, 1) == 0 || abs_num == 0 % check whether 'cell' is >= 10^3 or is an integer
                    val (k, 2) = sprintf ('%.0f', cell);
                elseif abs_num >= 10
                    val (k, 2) = sprintf ('%.2f', cell);
                elseif abs_num >= 1
                    val (k, 2) = sprintf ('%.3f', cell); 
                else
                    val (k, 2) = sprintf ('%.4f', cell);
                end
            else % else if val(k,2) is a string
                cell = val (k, 2);
            end

            % write 'cell' content on .xlsx file
            range = sprintf ('%s%d', num2xlcol(j), i); % get cell address
            writematrix (cell, xlsxName, 'Sheet', 1, 'Range', range, 'AutoFitWidth', false);

            j = j+1;
        else
            i = i+1;
            j=1;
        end
    end

    %% save data on .txt file
    line_size = round (1.5*max (sum (strlength (val), 2)));
    line_separator = char (45*ones (1, line_size+3));
    line_separator = ['+' line_separator '+'];
    table_separator = char (35*ones (1, line_size+5));
    fileID = fopen (txtName, 'w');

    fprintf (fileID, '%s\n%s\n', table_separator, line_separator);

    for k = 1:row
        if val (k, 1) ~= "#"
            % print text
            fprintf (fileID, '| %s:', val (k, 1));

            % print spaces
            space_size = line_size - sum (strlength (val (k, :)));
            space = char (32*ones (1, space_size));
            fprintf (fileID, space);

            % print values
            fprintf (fileID, '%s |\n', val (k, 2));

            % print line_separator
            fprintf (fileID, '%s\n', line_separator);
        else
            fprintf (fileID, '%s\n%s\n', table_separator, line_separator);
        end
    end

    fprintf (fileID, table_separator);
    fclose (fileID);

    %% type data on command prompt
    type (txtName)
    fprintf ('\n\n\n')
end


function xlcol_addr = num2xlcol (col_num)
    % col_num - positive integer greater than zero

    n=1;

    while col_num > 26*(26^n-1)/25
        n = n+1;
    end

    base_26 = zeros(1,n);
    tmp_var = -1+col_num-26*(26^(n-1)-1)/25;

    for k = 1:n
        divisor = 26^(n-k);
        remainder = mod(tmp_var,divisor);
        base_26(k) = 65+(tmp_var-remainder)/divisor;
        tmp_var = remainder;
    end

    xlcol_addr = char(base_26); % Character vector of xlcol address
end
