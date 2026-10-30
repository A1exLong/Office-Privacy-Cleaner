# Office Privacy Cleaner · Office 身份元数据清理器

本地控制台工具：把 Office 文件拖到 exe 上，生成同目录的 `_clean` 副本；已有同名输出在处理、校验成功后覆盖。**源文件始终保留**。程序不会联网或上传文档。

此工具清理已知文档属性和协作身份信息，**不是正文个人信息识别或完整匿名化工具**。正文中直接写出的姓名、邮箱、证件号，图片、宏、附件和业务数据里的身份信息不会自动识别。输出仍需按使用场景复核。

## 使用

将一个或多个文件拖到：

```text
dist/OfficePrivacyCleaner.exe
```

例如 `报告.docx` → `报告_clean.docx`。再次处理原件时覆盖原来的清理副本。处理 `报告_clean.docx` 时生成 `报告_clean_clean.docx`，避免覆盖本次输入。文件被占用或没有写权限时明确报错；不会改写源文件，也不会删除先前的有效输出。

也可双击 exe 后输入单个文件路径，或从终端运行：

```powershell
.\OfficePrivacyCleaner.exe "D:\资料\报告.docx" "D:\资料\数据.xlsx" --no-pause
```

正常任务返回 0；任一文件或报告写入失败返回 1，其余文件继续处理。exe 默认等待回车再关闭，以便查看结果；脚本模式不等待。退出码 0 代表已完成已声明的规则，不表示所有个人信息都已消失。

Python 源码运行：

```powershell
python -m pip install -r requirements.txt
python main.py "D:\资料\报告.docx"
```

额外传入 `--report` 会在输出旁写入 `.clean-report.json`，包含清理计数、范围说明及源/目标路径。报告本身含本机路径，不建议随清理后的文档一起外发。

## 格式支持

| 类型 | 支持格式 | 处理方式 |
| --- | --- | --- |
| Word | `.docx .docm .dotx .dotm` | 直接处理 OOXML 包，无需安装 Office |
| Excel | `.xlsx .xlsm .xltx .xltm` | 同上，保留公式、工作表和样式 |
| PowerPoint | `.pptx .pptm .potx .potm .ppsx .ppsm` | 同上，保留幻灯片和备注 |
| 旧版 Word | `.doc .dot` | 需 Windows + 本机 Microsoft Word，转为 `.docx` 或 `.docm` 后清理 |
| 旧版 Excel | `.xls .xlt` | 需 Microsoft Excel，转为 `.xlsx` 或 `.xlsm` 后清理 |
| 旧版 PowerPoint | `.ppt .pps .pot` | 需 Microsoft PowerPoint，转为 `.pptx` 或 `.pptm` 后清理 |

旧版文件转换后的扩展名会变化；模板和放映文件转换为普通文档或演示文稿。转换器使用独立 Office 实例、禁用宏，不连接用户已经打开的实例，完成后关闭自己创建的实例。Office 仍可能对损坏、受保护或特殊文件显示对话框；旧格式自动化不能保证在无人值守场景下始终不阻塞。

不支持 `.xlsb`、PDF、ODF、密码加密/IRM 文件。仅安装 WPS 不能保证旧版转换器可用。转换可能影响旧版特有功能，请核对结果。

## 清理范围

| 项目 | 行为 |
| --- | --- |
| 核心、扩展、自定义文档属性 | 删除属性部件及关联；包括作者、修改者、创建修改时间、公司、经理、标题等 |
| 文档缩略图、包内数字签名 | 删除；清理后的副本不保留原签名效力 |
| Word 批注 | 删除批注、作者部件及正文中的批注范围/引用标记 |
| Word 修订身份 | 作者/缩写改为 `Anonymous`，删除编辑时间和 rsid；保留插入、删除、移动等修订内容 |
| Excel 批注 | 删除传统/线程批注、人员、修订日志及批注 VML 形状；不删除其它 VML 表单形状 |
| Excel 个人视图与共享用户名 | 删除；不删除业务工作表、隐藏工作表或公式 |
| PowerPoint 批注身份 | 删除评论及作者/人员关联部件 |
| ZIP 包元信息 | 清除包注释、额外字段、原始文件时间，重建为统一时间戳 |

明确保留并报告的内容：修订中的历史文字、演讲者备注、宏、嵌入对象、自定义 XML、媒体、外部链接。它们可能包含个人信息。图片 EXIF 等媒体元数据目前不会递归清理。隐藏文本、隐藏工作表、幻灯片对象名称、页眉页脚、公式缓存、外部链接缓存等也不做身份识别。

删除文档属性可能让 `DOCPROPERTY` 等字段在下次更新后变为空值；已经渲染的字段文字保留。打开并重新保存文件时，Office 可能重新写入当前用户的元数据；需要时请在最终保存后重新运行工具。文件名、目录、NTFS 时间/权限/备用数据流、云端版本历史不属于文档包清理范围。

## 架构

```text
Office-Privacy-Cleaner/
├── main.py                    控制台启动入口
├── office_cleaner/
│   ├── models.py              报告、异常和处理器协议
│   ├── package.py             安全 ZIP/XML 读写、关系修复和校验
│   ├── rules.py               元数据、协作身份、保留风险清单规则
│   ├── ooxml.py               新版 Office 格式适配器
│   ├── legacy.py              可选 Microsoft Office 旧版转换器
│   ├── engine.py              注册表分发、处理流程与原子替换
│   └── cli.py                 参数、批处理和报告输出
└──  requirements.txt           运行依赖
```

新增格式：实现 `Cleaner` 协议，注册到 `DEFAULT_CLEANERS`，无需改动 CLI。新增 OOXML 清理策略：实现 `Rule.apply(package, report)`，注册到 `DEFAULT_RULES`，或给 `OOXMLCleaner(rules=...)` 注入规则集合。

处理全程不解压到任意路径；限制输入、解压总量、单部件大小与部件数，拒绝重复/非法路径及 DTD。XML 解析关闭外部实体、DTD 加载和网络访问。先校验原始包的关系，再修复删除部件的引用，输出重新打开校验后才原子替换目标。当前限制为输入 256 MiB、单部件 128 MiB、解压总量 1 GiB、20,000 部件、压缩比 1000 倍；不适合超大文件或完全不可信文件的多租户服务。

## 参考

- [Microsoft 文档检查器及其覆盖范围](https://support.microsoft.com/en-us/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks)
- [Word RemovePersonalInformation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.removepersonalinformation)
- [Word 修订接受涉及的结构](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)

本工具实现的是上表中声明的规则，不宣称等价于 Office 文档检查器的所有检查项。
