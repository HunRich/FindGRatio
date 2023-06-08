clc; 
axonDiameter = []; %creates open list to hold axon diameters
prompt1 = 'Input the file path for ROI zip file: '; %pathway needs to lead to ROI zip with ROI's
prompt2 = 'Input the file path for centroid excel file: '; %pathway needs to lead to csv file with centroid measurements, one row per ROI 
prompt3 = 'Input the pixels/micron scale factor: '; %Needs to be a number
ROIpath = input(prompt1, 's'); %assigns the path given to a variable
centPath = input(prompt2, 's'); %assigns the path given to a variable
pixToMicron = 39.2081; %input(prompt3); %assigns the pixel/micron scaling factor to a variable
    
counter = 0; %set up variables to count the number of input tries
counter2 = 0;
counter3 = 0;
while counter <= 2 %while loop to assure the input ROI path works, gives 3 retries
    try 
        axonROI = ReadImageJROI(ROIpath); %Takes the ROI's in the zip file into Matlab as an array
        break %goes out of the while loop if sucessful
    catch %if there is an error
        disp(newline); %makes a new line for readability
        disp('Please make sure the ROI pathway is correct. Ex.C:\Users\James\Documents\Research\RoiSet.zip');
        ROIpath = input(prompt1, 's'); %asks again for pathway
        counter = counter + 1;
    end
end
while counter2 <= 2 && counter ~= 3 %while loop to assure the inputed centroid path works, gives 3 retries
    try
        centroids = UploadExceltoMatlab(centPath); %Takes the list of all of the centroids into Matlab
        break %goes out of the while loop if sucessful
    catch %if there is an error
        disp(newline);
        disp('Please make sure the centroid pathway is correct. Ex. C:\Users\James\Documents\Research\Results.csv');
        centPath = input(prompt2, 's'); %asks again for pathway
        counter2 = counter2 + 1;
    end
end
while ~isreal(pixToMicron) == 1 || pixToMicron <= 0 %while loop to make sure the entered number is real and positive
        disp(newline);
        disp('Please make sure the pixel/ micron scaling factor is correct. Ex. 243.7770');
        pixToMicron = input(prompt3);
        counter3 = counter3 + 1;
        if counter3 > 2
            break
        end
end
%end program after 3 unsuccessful retries
if counter == 3 || counter2 == 3 || counter3 == 3
    disp(newline);
    disp('Please look up the pathway or scaling factor and try again.');
    return
end

axonNumber = [];
for i = 1:length(axonROI) %for loop that iterates over the list of ROI's
    outline = axonROI{i}; %Takes the first array and assigns it as the inner myelin tracing
    outline.mnCoordinates; %Takes the pixel coordinates for this space'
    centr = centroids{i, 2:3}; %Takes the centroid for this axon
    %changes the coordinates from pixels to microns using the given scaling factor and then
    %makes a list of all of the distances from the centroid to each pixel
    outlineRad = sqrt(((outline.mnCoordinates(:,1)/pixToMicron)-centr(1)).^2 + ((outline.mnCoordinates(:,2)/pixToMicron)-centr(2)).^2);
    %takes the avg of the radians
    outlineAvgRad = mean(outlineRad);
    outlineAvgDiameter = outlineAvgRad * 2; %Multiplies the avg radius by 2 to find the avg diameter
    axonDiameter (end+1) = outlineAvgDiameter; %adds the axon diameter to a list
    axonName = convertCharsToStrings(outline.strName); %takes the name of the ROI from Fiji Data
    axonNumber (end+1) = axonName; %adds that name to a list
end

disp(newline);
axonData = [axonNumber; axonDiameter]; %makes the two a matrix for displaying as a table
dataTable = array2table(axonData, 'RowNames',{'Axon Number', 'Axon Diameter (microns)'}); %makes the table
disp(dataTable); %prints the data with row labels

disp(newline);
prompt4 = 'Would you like to save the data to a spreadsheet? [Y/N]: ';
prompt5 = 'Please Enter the file name, including .xls at the end: ';
ques =  input(prompt4, 's'); %ask for answer on whether to create a spreadsheet
counter4 = 0;
while counter4 <= 2 %loop to check input for questions about the spreadsheet
    if strcmp(ques,'Y') == 1 || strcmp(ques,'y') == 1
        filename =  input(prompt5, 's'); %ask for what the file should be named
        try
            writetable(dataTable, filename); %makes a file ( any .xls, .csv, or .txt depending on extension used) 
                                             %with the table and saves it to the matlab pathway
            break
        catch
            disp(newline);
            disp('Please enter a file name with the correct format. Ex. ImageData.xls')
            counter4 = counter4 + 1;
        end
    elseif strcmp(ques,'N') == 1 || strcmp(ques,'n') == 1
        break
    else
        disp(newline);
        disp('Please type in Y or N to answer the question.');
        counter4 = counter4 + 1;
    end
end
