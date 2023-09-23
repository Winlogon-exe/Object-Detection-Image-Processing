#include <opencv2/opencv.hpp>
#include <iostream>
#include <xlsxwriter.h>
#include <string>
#include <filesystem>

namespace fs = std::filesystem;

//Функция для создания таблицы Excel
//Error форматирование??
void createExcelTable(const std::string& filename, double truePositiveRate, double falsePositiveRate) {

    lxw_workbook* workbook = workbook_new(filename.c_str());
    lxw_worksheet* worksheet = workbook_add_worksheet(workbook, NULL);

    //заголовки для столбцов
    worksheet_write_string(worksheet, 0, 0, "Метрика", NULL);
    worksheet_write_string(worksheet, 0, 1, "Значение", NULL);

    //Запись процентов верных и ложных обнаружений
    worksheet_write_string(worksheet, 1, 0, "True Positive Rate (%)", NULL);
    worksheet_write_number(worksheet, 1, 1, truePositiveRate, NULL);

    worksheet_write_string(worksheet, 2, 0, "False Positive Rate (%)", NULL);
    worksheet_write_number(worksheet, 2, 1, falsePositiveRate, NULL);

    //освобождение ресурсов
    int result = workbook_close(workbook);
     if (result == LXW_NO_ERROR) {
        std::cout<<("Таблица Excel создана успешно.\n");
    }
    else {
         std::cerr << (stderr, "Ошибка: Не удалось создать таблицу Excel.\n"); 
    }
}

//Функция для вычисления процента верных и ложных обнаружений
void calculateDetectionAccuracy(const std::vector<std::vector<cv::Point>>& contours,int expectedItemCount, double minContourSize, double maxContourSize) {
    int actualItemCount = contours.size();

    int truePositives = 0;
    int falsePositives = 0;

    for (const std::vector<cv::Point>& contour : contours) {
        double contourArea = cv::contourArea(contour);
        if (contourArea >= minContourSize && contourArea <= maxContourSize) {
            truePositives++;
        }
        else {
            falsePositives++;
        }
    }

    int falseNegatives = expectedItemCount - truePositives;

    double truePositiveRate = static_cast<double>(truePositives) / expectedItemCount * 100.0;
    double falsePositiveRate = static_cast<double>(falsePositives) / expectedItemCount * 100.0;

    std::cout << "Expected Items: " << expectedItemCount << std::endl;
    std::cout << "Actual Items Detected: " << actualItemCount << std::endl;
    std::cout << "True Positives: " << truePositives << std::endl;
    std::cout << "False Positives: " << falsePositives << std::endl;
    std::cout << "False Negatives: " << falseNegatives << std::endl;
    std::cout << "True Positive Rate (%): " << truePositiveRate << std::endl;
    std::cout << "False Positive Rate (%): " << falsePositiveRate << std::endl;

    createExcelTable("detection_accuracy.xlsx", truePositiveRate, falsePositiveRate);
}

// Функция для обработки изображения
cv::Mat processImage(const cv::Mat& inputImage) {
    cv::Mat inputCopy = inputImage.clone();

    //Преобразование изображения в градации серого
    cv::Mat gray;
    cv::cvtColor(inputCopy, gray, cv::COLOR_BGR2GRAY);

    //Размытие изображения, чтобы удалить шумы
    cv::GaussianBlur(gray, gray, cv::Size(5, 5), 0);

    //Применение оператора Кэнни для поиска границ
    cv::Mat edges;
    cv::Canny(gray, edges, 50, 100);

    //Закрытие границ объектов
    cv::Mat kernel = cv::getStructuringElement(cv::MORPH_RECT, cv::Size(3, 3));
    cv::morphologyEx(edges, edges, cv::MORPH_CLOSE, kernel);

    //Найдем внешние контуры
    std::vector<std::vector<cv::Point>> contours;
    cv::findContours(edges.clone(), contours, cv::RETR_EXTERNAL, cv::CHAIN_APPROX_SIMPLE);

    //Входные параметры коробки
    double minContourSize = 1000.0;
    double maxContourSize = 100000.0;

    for (const std::vector<cv::Point>& contour : contours) {
        double contourArea = cv::contourArea(contour);
        if (contourArea >= minContourSize && contourArea <= maxContourSize) {
            cv::RotatedRect rect = cv::minAreaRect(contour);

            cv::Point2f box[4];
            rect.points(box); //Получение вершин прямоугольника
            for (int j = 0; j < 4; ++j) {
                cv::line(inputCopy, box[j], box[(j + 1) % 4], cv::Scalar(0, 0, 255), 2); //Рисование прямоугольника 
            }
        }
    }
    int expectedItemCount = 3; //ожидаемое количество объектов на изображении
    calculateDetectionAccuracy(contours, expectedItemCount, minContourSize, maxContourSize);

    return inputCopy;
}

//Для сохранения изображения
bool saveImage(const std::string& outputPath, const cv::Mat& image) {
    static int imageCount = 0;
    std::string filename = outputPath + "/output" + std::to_string(imageCount++) + ".jpg";

    if (cv::imwrite(filename, image)) {
        std::cout << "Image saved successfully: " << filename << std::endl;
        return true;
    }
    else {
        std::cerr << "Error: Could not save image." << std::endl;
        return false;
    }
}


int main() {
    setlocale(LC_ALL, "RUS");
    const std::string inputFolder = "D:/zadanie/"; //Папка с изображениями
    const std::string outputFolder = "output_images"; //Папка для  обработанных изображений

    //папку для сохранения обработанных изображений, если она не существует
    if (!fs::exists(outputFolder)) { 
        fs::create_directory(outputFolder); 
    }

    //по всем изображениям в папке
    for (const auto& entry : fs::directory_iterator(inputFolder)) {
        std::string imagePath = entry.path().string();
        cv::Mat inputImage = cv::imread(imagePath);

        if (!inputImage.empty()) {
            cv::Mat resultImage = processImage(inputImage);
            saveImage(outputFolder, resultImage);
        }
        else {
            std::cerr << "Error: Could not read image" << imagePath << std::endl;
        }
    }

    cv::destroyAllWindows();
    return 0;
}
