# PP5AutoUITests

PP5AutoUITests 是一個自動化UI測試項目，使用C#和Appium進行測試。該項目旨在測試PowerPro5應用程序的各種功能。

## 目錄結構

- **根目錄**：
  - `app.config` 和 `app.config.bak`：應用程序的配置文件。
  - `AppiumServerManager.cs`：管理Appium服務器的文件。
  - `Assertor.cs`：包含自定義斷言方法。
  - `PP5AutoUITests.csproj` 和 `PP5AutoUITests.csproj.bak`：項目文件和備份文件。

- **Base目錄**：
  - `DriverBase.cs`：基礎驅動類，用於初始化和管理WebDriver實例。

- **Converter目錄**：
  - 包含數據轉換相關的類和方法。

- **ElementAction.cs** 和 **ElementFinder.cs**：
  - `ElementAction.cs`：包含對網頁元素進行操作的方法。
  - `ElementFinder.cs`：包含查找網頁元素的方法。

- **ElementFinderTests.cs**：
  - 測試 `ElementFinder` 類的方法。

- **Enum目錄**：
  - 包含枚舉類，用於定義一組命名常數。

- **Event目錄**：
  - 包含事件處理相關的類和方法。

- **Exception目錄**：
  - 包含自定義異常類。

- **Extension目錄**：
  - 包含擴展方法。

- **Factory目錄**：
  - 包含工廠模式相關的類。

- **Helper目錄**：
  - 包含幫助類和方法。

- **Interface目錄**：
  - 包含接口定義。

- **Model目錄**：
  - 包含數據模型類。

- **Module_TestCases目錄**：
  - 包含測試用例。

- **bin目錄**和**obj目錄**：
  - `bin` 目錄包含編譯後的二進制文件。
  - `obj` 目錄包含中間文件和編譯過程中的臨時文件。

- **packages.config**：
  - NuGet包管理配置文件。

- **PerformAction.cs**、**PowerPro5Config.cs**、**PowerPro5ConfigTests.cs**：
  - `PerformAction.cs`：包含執行操作的方法。
  - `PowerPro5Config.cs`：包含應用程序的配置設置。
  - `PowerPro5ConfigTests.cs`：包含測試 `PowerPro5Config` 類的方法。

- **Properties目錄**：
  - 包含程序集信息文件。

- **XML文檔**：
  - 包含類和方法的文檔說明。

- **TestResults目錄**：
  - 包含測試結果文件。

## 安裝

1. 克隆此存儲庫到本地機器：
    ```sh
    git clone https://github.com/yourusername/PP5AutoUITests.git
    ```

2. 打開項目文件 [PP5AutoUITests.csproj](http://_vscodecontentref_/0)。

3. 使用NuGet恢復依賴項：
    ```sh
    nuget restore
    ```

## 使用

1. 配置 [app.config](http://_vscodecontentref_/1) 文件以設置應用程序的配置。

2. 使用Visual Studio或其他IDE運行測試。

3. 查看 [TestResults](http://_vscodecontentref_/2) 目錄中的測試結果。

## 貢獻

歡迎貢獻！請遵循以下步驟：

1. Fork 此存儲庫。
2. 創建一個新的分支 (`git checkout -b feature/your-feature`)。
3. 提交您的更改 (`git commit -am 'Add some feature'`)。
4. 推送到分支 (`git push origin feature/your-feature`)。
5. 創建一個新的Pull Request。

## 許可

此項目使用 MIT 許可證。