# ExcelMerge AI 安装说明

这份文档供 AI 代理或自动化工具在用户授权后安装 ExcelMerge，并为 Git 配置表格文件的三方合并。当前发布包通常包含：

    bin/ExcelMerge.Cli.exe
    bin/ExcelMerge.Desktop.exe
    docs/AI-INSTALL.md
    docs/ai-git-merge-driver.md
    docs/user-guide.zh-CN.md

本文以 Windows PowerShell 为例。非 Windows 包使用 bin/ExcelMerge.Cli 和 bin/ExcelMerge.Desktop 作为入口文件名，并将路径分隔符替换为当前系统格式。

项目源码地址：[https://github.com/ksgfk/ExcelMerge](https://github.com/ksgfk/ExcelMerge)。

如果安装、启动、Git driver、文件解析或合并行为出现问题，AI 应先读取本包中的文档和清单；如果文档不足，再阅读上述源码仓库中的当前实现、测试和项目配置，以源码为准确定问题。若包来自特定提交，应优先检查与该包对应的提交。

## 安装前检查

1. 先确认用户要使用哪种范围：
   - **单个工程**：只影响当前 Git 仓库，推荐默认使用。
   - **全局**：影响当前用户的所有 Git 仓库，只有在用户明确要求时使用。
2. 确认包来自可信来源，并确认当前包的运行平台与目标机器匹配。
3. 不要覆盖用户已有的 merge.excelmerge.driver 配置，除非用户明确要求更新。
4. 不要在没有确认的情况下修改全局 Git 配置、全局 attributes 文件或其他工程。

## 安装 CLI 和 Desktop

将包内 CLI 和 Desktop 复制到稳定位置。不要把 Git 配置指向临时解压目录，因为删除或移动该目录会使后续 Git 合并失败。

    $packageRoot = (Resolve-Path "C:/Downloads/ExcelMerge-win-x64").Path
    $installRoot = Join-Path $env:LOCALAPPDATA "ExcelMerge"
    New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
    Copy-Item -Path (Join-Path $packageRoot "bin/*") -Destination $installRoot -Recurse -Force
    $cli = Join-Path $installRoot "ExcelMerge.Cli.exe"
    $desktop = Join-Path $installRoot "ExcelMerge.Desktop.exe"
    & $cli --version
    if (-not (Test-Path -LiteralPath $desktop -PathType Leaf)) {
        throw "The packaged Desktop app was not found."
    }

只有 CLI 的 --version 成功且 Desktop 文件存在后，才继续写入 Git 配置。不要删除 desktop 目录中的 av_libglesv2.dll、libHarfBuzzSharp.dll、libSkiaSharp.dll 等运行时文件。需要图形界面时启动 $desktop。保留 $installRoot 及其内容；Git driver 会长期调用其中的 CLI。

## 仅配置当前工程

这是默认推荐的方式。它只影响当前仓库，不改变用户的其他项目。

在目标仓库根目录执行，或把 $repoRoot 换成目标仓库路径：

    $repoRoot = (Resolve-Path "C:/work/my-repository").Path
    $cli = Join-Path $env:LOCALAPPDATA "ExcelMerge/ExcelMerge.Cli.exe"
    $driver = '"{0}" merge-driver %O %A %B %L %P' -f ($cli -replace '\\', '/')

    $attributesPath = Join-Path $repoRoot ".gitattributes"
    $rules = @(
        "*.xlsx merge=excelmerge"
        "*.csv  merge=excelmerge"
        "*.tsv  merge=excelmerge"
    )

    if (-not (Test-Path -LiteralPath $attributesPath)) {
        Set-Content -LiteralPath $attributesPath -Value $rules -Encoding UTF8
    } else {
        $existing = Get-Content -LiteralPath $attributesPath
        $missing = $rules | Where-Object { $_ -notin $existing }
        if ($missing) {
            Add-Content -LiteralPath $attributesPath -Value $missing -Encoding UTF8
        }
    }

    $currentDriver = git -C $repoRoot config --local --get merge.excelmerge.driver
    if (-not [string]::IsNullOrWhiteSpace($currentDriver) -and $currentDriver -ne $driver) {
        throw "This repository already has a different merge.excelmerge.driver. Ask before replacing it."
    }
    if ([string]::IsNullOrWhiteSpace($currentDriver)) {
        git -C $repoRoot config --local merge.excelmerge.name "ExcelMerge 三方工作簿合并"
        git -C $repoRoot config --local merge.excelmerge.driver $driver
    }

    git -C $repoRoot config --show-origin --get merge.excelmerge.driver
    git -C $repoRoot check-attr merge -- report.xlsx data.csv data.tsv

将 .gitattributes 提交到工程，这样其他协作者能获得文件类型关联；但每台机器仍需配置自己的 CLI 路径。若工程已经有 .gitattributes，只追加缺少的规则，不要覆盖原有规则。

## 全局配置

全局配置会影响当前用户的所有 Git 仓库。除了定义 driver，还必须配置全局 attributes 文件，否则 Git 不会自动把 .xlsx、.csv、.tsv 交给该 driver。

    $cli = Join-Path $env:LOCALAPPDATA "ExcelMerge/ExcelMerge.Cli.exe"
    $driver = '"{0}" merge-driver %O %A %B %L %P' -f ($cli -replace '\\', '/')
    $globalAttributes = git config --global --get core.attributesFile
    if ([string]::IsNullOrWhiteSpace($globalAttributes)) {
        $globalAttributes = Join-Path $env:USERPROFILE ".config/git/attributes"
        git config --global core.attributesFile ($globalAttributes -replace '\\', '/')
    } else {
        $globalAttributes = $globalAttributes.Trim()
    }

    New-Item -ItemType Directory -Path (Split-Path $globalAttributes) -Force | Out-Null
    if (-not (Test-Path -LiteralPath $globalAttributes)) {
        Set-Content -LiteralPath $globalAttributes -Value @(
            "*.xlsx merge=excelmerge"
            "*.csv  merge=excelmerge"
            "*.tsv  merge=excelmerge"
        ) -Encoding UTF8
    } else {
        $existing = Get-Content -LiteralPath $globalAttributes
        $rules = @(
            "*.xlsx merge=excelmerge"
            "*.csv  merge=excelmerge"
            "*.tsv  merge=excelmerge"
        )
        $missing = $rules | Where-Object { $_ -notin $existing }
        if ($missing) {
            Add-Content -LiteralPath $globalAttributes -Value $missing -Encoding UTF8
        }
    }

    $currentDriver = git config --global --get merge.excelmerge.driver
    if (-not [string]::IsNullOrWhiteSpace($currentDriver) -and $currentDriver -ne $driver) {
        throw "The global Git configuration already has a different merge.excelmerge.driver. Ask before replacing it."
    }
    if ([string]::IsNullOrWhiteSpace($currentDriver)) {
        git config --global merge.excelmerge.name "ExcelMerge 三方工作簿合并"
        git config --global merge.excelmerge.driver $driver
    }

    git config --show-origin --global --get core.attributesFile
    git config --show-origin --global --get merge.excelmerge.driver

全局 attributes 文件是用户级配置，不应提交到任意工程。全局模式可能改变第三方仓库或不适合使用 ExcelMerge 的项目；如用户只想影响一个工程，应改用上一节的单工程配置。

## 验证配置

在目标仓库中验证：

    git check-attr -a -- report.xlsx data.csv data.tsv
    git config --get merge.excelmerge.driver

预期结果应包含 merge: excelmerge，driver 命令中必须完整保留 %O %A %B %L %P。其中 %O 是 BASE，%A 是当前分支/OURS，%B 是远端/THEIRS，%L 是 Git 冲突标记大小，%P 是仓库内原始路径。

可以用三份临时的同格式文件直接验证：

    Copy-Item base.xlsx test-base.xlsx
    Copy-Item local.xlsx test-local.xlsx
    Copy-Item remote.xlsx test-remote.xlsx
    & $cli merge-driver test-base.xlsx test-local.xlsx test-remote.xlsx 7 data/report.xlsx

成功返回 0 时，test-local.xlsx 会被替换为合并结果。返回 1 时不要执行 git add；它可能表示未解决冲突，也可能表示操作失败。

## Git 冲突解决

自动合并失败后：

1. 保留 Git 的未合并状态，执行 git status。
2. 获取同一个路径的 BASE、LOCAL 和 REMOTE 三个版本，使用二进制安全方式写入临时文件。
3. 运行桌面程序的三方合并：

       dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- merge BASE LOCAL REMOTE RESULT

   发布包同时包含 CLI 和 Desktop。可以直接启动已安装的 $desktop；如果包与当前系统平台不匹配，再使用源码构建或安装对应平台的发布包。
4. 在冲突列表中选择“使用本地”“使用远端”“保留两者”或适用的自定义值。
5. 保存到原始冲突路径，确认工作簿有效后执行 git add -- <path>。

“保留两者”是行级决策，不是普通单元格取值。没有解决全部冲突时，不要提交、不执行 git add，也不要用 git checkout --ours 或 git checkout --theirs 来伪造合并结果。

详细的 Git 占位符、三方语义、冲突导出和自动化约束见 [ai-git-merge-driver.md](ai-git-merge-driver.md)。面向用户的操作说明见 [user-guide.zh-CN.md](user-guide.zh-CN.md)。

## 卸载配置

单工程：

    git -C C:/work/my-repository config --local --unset merge.excelmerge.driver
    git -C C:/work/my-repository config --local --unset merge.excelmerge.name

然后从该工程的 .gitattributes 删除 ExcelMerge 规则。不要删除其他 attributes 规则。

全局：

    git config --global --unset merge.excelmerge.driver
    git config --global --unset merge.excelmerge.name
    # 仅在安装前没有 core.attributesFile 配置时执行
    git config --global --unset core.attributesFile

然后从全局 attributes 文件删除 ExcelMerge 规则。如果该文件没有其他规则，可以由用户手动删除它。卸载 driver 前先确认当前值仍是本次安装的 ExcelMerge driver；如果用户后来改成了其他 driver，不要执行 unset。只有在安装前不存在 core.attributesFile 配置时，才执行上面的 unset 命令；否则要保留原有配置。最后，在确认没有其他程序使用 CLI 后，再删除 %LOCALAPPDATA%/ExcelMerge。
