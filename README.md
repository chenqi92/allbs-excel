# allbs-excel

[![License](https://img.shields.io/badge/license-Apache%202-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)
[![Maven Central](https://img.shields.io/maven-central/v/cn.allbs/allbs-excel.svg)](https://search.maven.org/artifact/cn.allbs/allbs-excel)

基于 [Apache Fesod (Incubating)](https://fesod.apache.org/) 的 Spring Boot Excel 导入导出增强工具，通过注解即可实现 Excel 的导入导出功能。

## ✨ 特性

- 🚀 **简单易用**: 通过注解即可实现 Excel 导入导出
- 📝 **功能丰富**: 支持单/多 Sheet、模板导出、数据验证等
- 🎨 **灵活定制**: 支持自定义转换器、样式处理器
- 🔄 **字典转换**: 支持字典值与标签的自动转换
- 🔐 **数据脱敏**: 支持手机号、身份证等敏感数据脱敏
- 🌍 **国际化支持**: Excel 表头支持国际化
- 🔒 **数据验证**: 导入时自动进行数据校验
- 📊 **空数据导出**: 支持导出只有表头的空 Excel
- 🔀 **合并单元格**: 支持同值自动合并，支持依赖关系合并
- 📈 **进度回调**: 支持实时监听导出进度，适用于大数据量导出
- 🆕 **嵌套对象导出**: 支持从嵌套对象、集合、Map 中提取字段值
- 🆕 **对象自动展开**: 自动展开嵌套对象的所有字段
- 🆕 **List 展开**: 将 List 集合展开为多行，自动合并单元格
- 🆕 **条件样式**: 根据单元格值自动应用不同样式（颜色、字体等）
- 🆕 **动态表头**: 根据数据动态生成表头列，适用于自定义字段场景
- 🆕 **嵌套对象导入**: 导入时自动创建并填充嵌套对象
- 🆕 **List 聚合导入**: 将多行数据聚合回包含 List 的对象
- 🆕 **数据验证**: Excel 列添加数据验证规则（下拉列表、数值范围、日期等）
- 🆕 **多 Sheet 关联**: 主表和关联数据自动导出到不同 Sheet 并建立关联
~~- 🆕 **Excel 公式**: 支持在导出时自动添加 Excel 公式（SUM、AVERAGE、自定义公式等）~~
- 🆕 **冻结窗格**: 支持冻结指定行和列，方便查看大表格数据
- 🆕 **条件格式**: 高级条件格式，支持数据条、色阶、图标集等
- 🆕 **批注**: 为单元格添加批注说明
- 🆕 **图片导出**: 支持将图片（URL、本地路径、字节数组）嵌入到 Excel 单元格中
- 🆕 **Excel 加密**: 支持密码保护 Excel 文件（AES-256 加密）
- 🆕 **水印**: 为 Excel 添加水印保护（支持自定义文本、颜色、透明度、旋转角度）
- 🆕 **图表导出**: 在 Excel 中自动生成图表（折线图、柱状图、饼图、面积图、散点图等）
- ⚡ **高性能**: 基于 Apache Fesod 2.0.2-incubating，性能优异
- 🔄 **版本兼容**: 同时支持 Spring Boot 2.x 和 3.x

## 📦 依赖要求

- JDK 11+
- Spring Boot 2.7+ 或 3.x

## 🚀 快速开始

### 1. 添加依赖

```xml
<dependency>
    <groupId>cn.allbs</groupId>
    <artifactId>allbs-excel</artifactId>
    <version>3.2.0</version>
</dependency>
```

**注意**: 本库同时支持 Spring Boot 2.x 和 3.x，无需额外配置。

### 2. 创建实体类

```java
@Data
public class UserDTO {
    @ExcelProperty(value = "用户ID", index = 0)
    private Long id;

    @ExcelProperty(value = "用户名", index = 1)
    private String username;

    @ExcelProperty(value = "邮箱", index = 2)
    @Email(message = "邮箱格式不正确")
    private String email;

    @ExcelProperty(value = "创建时间", index = 3)
    @DateTimeFormat("yyyy-MM-dd HH:mm:ss")
    private LocalDateTime createTime;
}
```

### 3. 导出 Excel

```java
@RestController
@RequestMapping("/user")
public class UserController {

    @GetMapping("/export")
    @ExportExcel(
        name = "用户列表",
        sheets = @Sheet(sheetName = "用户信息")
    )
    public List<UserDTO> exportUsers() {
        return userService.findAll();
    }
}
```

访问 `/user/export` 即可下载 Excel 文件。

### 4. 导入 Excel

```java
@PostMapping("/import")
public ResponseEntity<?> importUsers(@ImportExcel List<UserDTO> users) {
    userService.batchSave(users);
    return ResponseEntity.ok("导入成功");
}
```

## 📖 详细使用说明

### 一、导出功能

#### 1.1 基本导出

最简单的导出方式，返回 List 即可：

```java
@GetMapping("/export")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(sheetName = "用户信息")
)
public List<UserDTO> exportUsers() {
    return userService.findAll();
}
```

#### 1.2 空数据导出（带表头）

当数据为空时，也可以导出只有表头的 Excel。需要在 `@Sheet` 注解中指定 `clazz` 属性：

```java
@GetMapping("/export-empty")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(
        sheetName = "用户信息",
        clazz = UserDTO.class  // ⭐ 关键：指定数据类型用于生成表头
    )
)
public List<UserDTO> exportEmpty() {
    return Collections.emptyList();  // 会导出带表头的空 Excel
}
```

**说明**:
- 如果指定了 `clazz`，空数据时会根据该类型生成表头
- 如果未指定 `clazz`，空数据时只会创建一个空的 sheet（无表头）

#### 1.3 列顺序控制

使用 `@ExcelProperty` 的 `index` 属性可以控制列的顺序，**支持非连续的索引值**：

```java
@Data
public class UserDTO {
    @ExcelProperty(value = "姓名", index = 1)
    private String name;

    @ExcelProperty(value = "年龄", index = 2)
    private Integer age;

    @ExcelProperty(value = "地址", index = 7)
    private String address;

    @ExcelProperty(value = "备注", index = 11)
    private String remark;
}
```

**导出结果**：
- 第 1 列（B列）：姓名
- 第 2 列（C列）：年龄
- 第 7 列（H列）：地址
- 第 11 列（L列）：备注
- 其他列（A、D、E、F、G、I、J、K）：空列

**说明**：
- `index` 不需要从 0 开始，也不需要连续
- 列的顺序完全由 `index` 的值决定
- 未指定 `index` 的字段会按照字段定义顺序排列

#### 1.4 只导出有注解的字段

默认情况下，Apache Fesod 会导出所有字段。如果只想导出标注了 `@ExcelProperty` 的字段，可以使用 `onlyExcelProperty` 配置：

**方式一：在 @ExportExcel 中全局配置**

```java
@GetMapping("/export")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(sheetName = "用户信息"),
    onlyExcelProperty = true  // ⭐ 只导出有 @ExcelProperty 注解的字段
)
public List<UserDTO> exportUsers() {
    return userService.findAll();
}
```

**方式二：在 @Sheet 中单独配置**

```java
@GetMapping("/export")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(
        sheetName = "用户信息",
        onlyExcelProperty = true  // ⭐ Sheet 级别配置，优先级更高
    )
)
public List<UserDTO> exportUsers() {
    return userService.findAll();
}
```

**实体类示例**：

```java
@Data
public class UserDTO {
    @ExcelProperty("用户ID")
    private Long id;

    @ExcelProperty("用户名")
    private String username;

    // 这个字段不会被导出（没有 @ExcelProperty 注解）
    private String password;

    // 这个字段不会被导出（没有 @ExcelProperty 注解）
    private String internalCode;
}
```

**说明**：
- `onlyExcelProperty = true` 时，只导出有 `@ExcelProperty` 注解的字段
- `onlyExcelProperty = false`（默认）时，导出所有字段
- Sheet 级别的配置优先级高于 ExportExcel 级别
- 等同于在实体类上添加 `@ExcelIgnoreUnannotated` 注解

#### 1.5 多 Sheet 导出

导出多个 Sheet 时，返回 `List<List<?>>` 类型，每个内层 List 对应一个 Sheet：

```java
@GetMapping("/export-multi")
@ExportExcel(
    name = "综合报表",
    sheets = {
        @Sheet(sheetName = "用户信息", clazz = UserDTO.class),
        @Sheet(sheetName = "订单信息", clazz = OrderDTO.class)
    }
)
public List<List<?>> exportMultiSheet() {
    List<UserDTO> users = userService.findAll();
    List<OrderDTO> orders = orderService.findAll();
    return Arrays.asList(users, orders);
}
```

**多 Sheet 空数据导出**:

```java
@GetMapping("/export-multi-empty")
@ExportExcel(
    name = "综合报表",
    sheets = {
        @Sheet(sheetName = "用户信息", clazz = UserDTO.class),
        @Sheet(sheetName = "订单信息", clazz = OrderDTO.class)
    }
)
public List<List<?>> exportMultiEmpty() {
    return Arrays.asList(
        Collections.emptyList(),  // 空用户数据，但有表头
        Collections.emptyList()   // 空订单数据，但有表头
    );
}
```

#### 1.6 模板导出

使用预定义的 Excel 模板进行导出：

```java
@GetMapping("/export-template")
@ExportExcel(
    name = "用户报表",
    template = "user-template.xlsx",  // 模板文件放在 resources/excel/ 目录下
    sheets = @Sheet(sheetName = "用户信息")
)
public List<UserDTO> exportWithTemplate() {
    return userService.findAll();
}
```

**模板文件位置**: `src/main/resources/excel/user-template.xlsx`

#### 1.7 动态文件名

支持使用 SpEL 表达式动态生成文件名，提供了丰富的预定义变量和自定义函数。

**基本用法**：

```java
@GetMapping("/export-dynamic")
@ExportExcel(
    name = "用户列表-#{#date}",  // 使用方法参数
    sheets = @Sheet(sheetName = "用户信息")
)
public List<UserDTO> exportDynamic(@RequestParam String date) {
    return userService.findByDate(date);
}
```

**支持的功能**：

##### 1. 方法参数访问

```java
// 简单参数
@ExportExcel(name = "报表-#{#date}")
public List<UserDTO> export(@RequestParam String date) { ... }

// 多个参数
@ExportExcel(name = "#{#startDate}-#{#endDate}-报表")
public List<UserDTO> export(@RequestParam String startDate, @RequestParam String endDate) { ... }

// 对象属性
@ExportExcel(name = "#{#user.name}-#{#user.department}")
public List<UserDTO> export(@RequestBody UserDTO user) { ... }
```

##### 2. 预定义变量

| 变量 | 类型 | 说明 | 示例 |
|------|------|------|------|
| `#now` | LocalDateTime | 当前日期时间 | `报表-#{#now}` |
| `#today` | LocalDate | 当前日期 | `报表-#{#today}` |
| `#timestamp` | Long | 当前时间戳（毫秒） | `报表-#{#timestamp}` |
| `#uuid` | String | 随机 UUID | `报表-#{#uuid}` |

```java
// 使用当前日期
@ExportExcel(name = "报表-#{#today}")
public List<UserDTO> export() { ... }
// 输出：报表-2024-01-15.xlsx

// 使用时间戳
@ExportExcel(name = "报表-#{#timestamp}")
public List<UserDTO> export() { ... }
// 输出：报表-1705305600000.xlsx

// 使用 UUID
@ExportExcel(name = "报表-#{#uuid}")
public List<UserDTO> export() { ... }
// 输出：报表-550e8400-e29b-41d4-a716-446655440000.xlsx
```

##### 3. 自定义函数

| 函数 | 参数 | 说明 | 示例 |
|------|------|------|------|
| `#formatDate()` | LocalDate, String | 格式化日期 | `#{#formatDate(#today, 'yyyyMMdd')}` |
| `#formatDateTime()` | LocalDateTime, String | 格式化日期时间 | `#{#formatDateTime(#now, 'yyyyMMdd_HHmmss')}` |
| `#sanitize()` | String | 清理文件名非法字符 | `#{#sanitize(#filename)}` |
| `#timestamp()` | - | 获取时间戳 | `#{#timestamp()}` |

```java
// 格式化日期
@ExportExcel(name = "报表-#{#formatDate(#today, 'yyyyMMdd')}")
public List<UserDTO> export() { ... }
// 输出：报表-20240115.xlsx

// 格式化日期时间
@ExportExcel(name = "报表-#{#formatDateTime(#now, 'yyyyMMdd_HHmmss')}")
public List<UserDTO> export() { ... }
// 输出：报表-20240115_103000.xlsx

// 清理文件名
@ExportExcel(name = "#{#sanitize(#filename)}")
public List<UserDTO> export(@RequestParam String filename) { ... }
// 输入：用户/列表:2024  输出：用户_列表_2024.xlsx
```

##### 4. 静态方法调用

```java
// 调用 Java 静态方法
@ExportExcel(name = "报表-#{T(java.time.LocalDate).now()}")
public List<UserDTO> export() { ... }

// 格式化日期
@ExportExcel(name = "报表-#{T(java.time.LocalDate).now().format(T(java.time.format.DateTimeFormatter).ofPattern('yyyyMMdd'))}")
public List<UserDTO> export() { ... }

// 获取系统属性
@ExportExcel(name = "报表-#{T(System).getProperty('user.name')}")
public List<UserDTO> export() { ... }
```

##### 5. 字符串操作

```java
// 大小写转换
@ExportExcel(name = "#{#name.toUpperCase()}-报表")
public List<UserDTO> export(@RequestParam String name) { ... }

// 字符串拼接
@ExportExcel(name = "#{#prefix + '-' + #suffix}")
public List<UserDTO> export(@RequestParam String prefix, @RequestParam String suffix) { ... }

// 字符串截取
@ExportExcel(name = "#{#name.substring(0, 5)}")
public List<UserDTO> export(@RequestParam String name) { ... }

// 字符串替换
@ExportExcel(name = "#{#name.replace(' ', '_')}")
public List<UserDTO> export(@RequestParam String name) { ... }
```

##### 6. 条件表达式

```java
// 三元运算符
@ExportExcel(name = "#{#type == 'user' ? '用户列表' : '订单列表'}")
public List<?> export(@RequestParam String type) { ... }

// 空值处理
@ExportExcel(name = "#{#name != null ? #name : '默认报表'}")
public List<UserDTO> export(@RequestParam(required = false) String name) { ... }

// Elvis 操作符
@ExportExcel(name = "#{#name ?: '默认报表'}")
public List<UserDTO> export(@RequestParam(required = false) String name) { ... }
```

##### 7. 数学运算

```java
// 页码计算
@ExportExcel(name = "第#{#page + 1}页报表")
public List<UserDTO> export(@RequestParam int page) { ... }

// 数量计算
@ExportExcel(name = "总计#{#count * 2}条")
public List<UserDTO> export(@RequestParam int count) { ... }
```

##### 8. 集合操作

```java
// 集合大小
@ExportExcel(name = "#{#ids.size()}条数据")
public List<UserDTO> export(@RequestParam List<Long> ids) { ... }

// 集合访问
@ExportExcel(name = "#{#names[0]}-报表")
public List<UserDTO> export(@RequestParam List<String> names) { ... }

// 集合判空
@ExportExcel(name = "#{#ids.isEmpty() ? '空数据' : '有数据'}")
public List<UserDTO> export(@RequestParam List<Long> ids) { ... }
```

**完整示例**：

```java
@GetMapping("/export-advanced")
@ExportExcel(
    name = "#{#sanitize(#department)}-#{#formatDate(#today, 'yyyyMMdd')}-#{#type == 'all' ? '全部' : '部分'}",
    sheets = @Sheet(sheetName = "数据")
)
public List<UserDTO> exportAdvanced(
    @RequestParam String department,
    @RequestParam String type
) {
    return userService.findByDepartmentAndType(department, type);
}
// 输出示例：技术部-20240115-全部.xlsx
```

**注意事项**：
- SpEL 表达式必须包含 `#` 符号才会被解析
- 如果表达式解析失败，会使用原始字符串作为文件名
- 建议使用 `#sanitize()` 函数清理用户输入的文件名，避免非法字符

#### 1.8 自定义样式

可以自定义表头和内容的样式：

```java
@GetMapping("/export-styled")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(sheetName = "用户信息"),
    writeHandler = {CustomStyleHandler.class}  // 自定义样式处理器
)
public List<UserDTO> exportStyled() {
    return userService.findAll();
}
```

#### 1.9 国际化表头

支持根据当前语言环境自动切换表头：

```java
@GetMapping("/export-i18n")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(sheetName = "用户信息"),
    i18nHeader = true  // 启用国际化
)
public List<UserDTO> exportI18n() {
    return userService.findAll();
}
```

**配置国际化资源文件** (`messages.properties`):
```properties
user.id=User ID
user.username=Username
user.email=Email
user.createTime=Create Time
```

#### 1.10 合并单元格

支持自动合并相同值的单元格，适用于分组数据展示：

**方式一：全局配置**

```java
@GetMapping("/export-merge")
@ExportExcel(
    name = "部门员工列表",
    sheets = @Sheet(sheetName = "员工信息"),
    autoMerge = true  // ⭐ 启用自动合并
)
public List<EmployeeDTO> exportWithMerge() {
    return employeeService.findAll();
}
```

**方式二：Sheet 级别配置**

```java
@GetMapping("/export-merge")
@ExportExcel(
    name = "部门员工列表",
    sheets = @Sheet(
        sheetName = "员工信息",
        autoMerge = true  // ⭐ Sheet 级别配置，优先级更高
    )
)
public List<EmployeeDTO> exportWithMerge() {
    return employeeService.findAll();
}
```

**实体类配置**：

```java
@Data
public class EmployeeDTO {
    @ExcelProperty(value = "部门", index = 0)
    @ExcelMerge  // ⭐ 标记需要合并的字段
    private String department;

    @ExcelProperty(value = "姓名", index = 1)
    @ExcelMerge(dependOn = "department")  // ⭐ 依赖部门列，只有部门相同时才合并
    private String name;

    @ExcelProperty(value = "职位", index = 2)
    @ExcelMerge(dependOn = "name")  // ⭐ 依赖姓名列
    private String position;

    @ExcelProperty(value = "工资", index = 3)
    private BigDecimal salary;
}
```

**导出效果**：

| 部门 | 姓名 | 职位 | 工资 |
|------|------|------|------|
| 技术部 | 张三 | Java工程师 | 15000 |
| ↑ | ↑ | 前端工程师 | 12000 |
| ↑ | 李四 | Python工程师 | 14000 |
| 市场部 | 王五 | 市场专员 | 8000 |

**说明**：
- `@ExcelMerge`：标记需要合并的字段
- `dependOn`：指定依赖的字段，只有依赖字段的值相同时，当前字段才会合并
- `enabled`：是否启用合并（默认 true）
- `autoMerge` 配置必须设置为 `true` 才会生效
- Sheet 级别的 `autoMerge` 配置优先级高于 ExportExcel 级别

**注意事项**：
- 合并功能需要数据按照合并字段排序，否则可能出现非预期的合并效果
- 建议在查询数据时使用 `ORDER BY` 对需要合并的字段进行排序
- 当前版本的合并功能基于 Apache Fesod 2.0.2-incubating 实现

#### 1.11 导出进度回调

支持实时监听导出进度，适用于大数据量导出场景：

**第一步：实现进度监听器**

```java
@Component
public class MyProgressListener implements ExportProgressListener {

    @Override
    public void onStart(int totalRows, String sheetName) {
        System.out.println("开始导出: " + sheetName + ", 总行数: " + totalRows);
    }

    @Override
    public void onProgress(int currentRow, int totalRows, double percentage, String sheetName) {
        System.out.printf("导出进度: %d/%d (%.2f%%) - %s%n",
            currentRow, totalRows, percentage, sheetName);
    }

    @Override
    public void onComplete(int totalRows, String sheetName) {
        System.out.println("导出完成: " + sheetName + ", 总行数: " + totalRows);
    }

    @Override
    public void onError(Exception exception, String sheetName) {
        System.err.println("导出失败: " + sheetName + ", 错误: " + exception.getMessage());
    }
}
```

**第二步：使用 @ExportProgress 注解**

```java
@GetMapping("/export-with-progress")
@ExportExcel(
    name = "用户列表",
    sheets = @Sheet(sheetName = "用户信息")
)
@ExportProgress(
    listener = MyProgressListener.class,  // ⭐ 指定进度监听器
    interval = 100  // ⭐ 每 100 行触发一次进度回调
)
public List<UserDTO> exportWithProgress() {
    return userService.findAll();
}
```

**进度回调配置**：

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `listener` | Class | - | 进度监听器类（必填） |
| `interval` | int | 100 | 进度更新间隔（行数） |
| `enabled` | boolean | true | 是否启用进度回调 |

**高级用法：WebSocket 实时推送进度**

```java
@Component
public class WebSocketProgressListener implements ExportProgressListener {

    @Autowired
    private SimpMessagingTemplate messagingTemplate;

    @Override
    public void onProgress(int currentRow, int totalRows, double percentage, String sheetName) {
        // 通过 WebSocket 推送进度到前端
        Map<String, Object> progress = new HashMap<>();
        progress.put("currentRow", currentRow);
        progress.put("totalRows", totalRows);
        progress.put("percentage", percentage);
        progress.put("sheetName", sheetName);

        messagingTemplate.convertAndSend("/topic/export-progress", progress);
    }

    // ... 其他方法实现
}
```

**说明**：
- 进度监听器必须实现 `ExportProgressListener` 接口
- `interval` 设置为 1 表示每行都触发回调（可能影响性能）
- `interval` 设置为 0 表示只在开始和结束时触发回调
- 进度回调在每个 Sheet 独立触发
- 支持与 WebSocket、SSE 等技术结合实现实时进度推送

### 二、导入功能

#### 2.1 基本导入

使用 `@ImportExcel` 注解自动解析上传的 Excel 文件：

```java
@PostMapping("/import")
public ResponseEntity<?> importUsers(@ImportExcel List<UserDTO> users) {
    userService.batchSave(users);
    return ResponseEntity.ok("导入成功，共 " + users.size() + " 条数据");
}
```

**前端上传示例**:
```html
<form method="post" enctype="multipart/form-data" action="/user/import">
    <input type="file" name="file" accept=".xlsx,.xls"/>
    <button type="submit">导入</button>
</form>
```

#### 2.2 带验证的导入

导入时自动进行数据校验：

```java
@PostMapping("/import-validate")
public ResponseEntity<?> importWithValidation(
    @ImportExcel List<UserDTO> users,
    BindingResult bindingResult
) {
    if (bindingResult.hasErrors()) {
        // 处理验证错误
        List<String> errors = bindingResult.getAllErrors()
            .stream()
            .map(ObjectError::getDefaultMessage)
            .collect(Collectors.toList());
        return ResponseEntity.badRequest().body(errors);
    }

    userService.batchSave(users);
    return ResponseEntity.ok("导入成功");
}
```




**实体类验证注解**:
```java
@Data
public class UserDTO {
    @NotNull(message = "用户ID不能为空")
    @ExcelProperty("用户ID")
    private Long id;

    @NotBlank(message = "用户名不能为空")
    @Size(min = 2, max = 20, message = "用户名长度必须在2-20之间")
    @ExcelProperty("用户名")
    private String username;

    @Email(message = "邮箱格式不正确")
    @ExcelProperty("邮箱")
    private String email;
}
```

#### 2.3 自定义导入监听器

可以自定义导入逻辑，实现更复杂的业务处理：

```java
@PostMapping("/import-custom")
public ResponseEntity<?> importCustom(
    @ImportExcel(readListener = CustomReadListener.class) List<UserDTO> users
) {
    return ResponseEntity.ok("导入成功");
}
```

#### 2.4 指定上传字段名

默认情况下，前端上传字段名为 `file`，可以自定义：

```java
@PostMapping("/import-custom-field")
public ResponseEntity<?> importCustomField(
    @ImportExcel(fileName = "excelFile") List<UserDTO> users
) {
    userService.batchSave(users);
    return ResponseEntity.ok("导入成功");
}
```

**前端上传**:
```html
<input type="file" name="excelFile" accept=".xlsx,.xls"/>
```

#### 2.5 跳过空行

导入时可以选择是否跳过空行：

```java
@PostMapping("/import-skip-empty")
public ResponseEntity<?> importSkipEmpty(
    @ImportExcel(ignoreEmptyRow = true) List<UserDTO> users
) {
    userService.batchSave(users);
    return ResponseEntity.ok("导入成功");
}
```

### 三、高级功能

#### 3.1 自定义转换器

对于特殊的数据类型，可以自定义转换器：

```java
@Data
public class UserDTO {
    @ExcelProperty(value = "状态", converter = StatusConverter.class)
    private Integer status;
}
```

**转换器实现**:

```java
public class StatusConverter implements Converter<Integer> {
    @Override
    public Integer convertToJavaData(ReadCellData<?> cellData,
                                      ExcelContentProperty contentProperty,
                                      GlobalConfiguration globalConfiguration) {
        String stringValue = cellData.getStringValue();
        if ("启用".equals(stringValue)) {
            return 1;
        } else if ("禁用".equals(stringValue)) {
            return 0;
        }
        return null;
    }

    @Override
    public WriteCellData<?> convertToExcelData(Integer value,
                                                 ExcelContentProperty contentProperty,
                                                 GlobalConfiguration globalConfiguration) {
        if (value == 1) {
            return new WriteCellData<>("启用");
        } else if (value == 0) {
            return new WriteCellData<>("禁用");
        }
        return new WriteCellData<>("");
    }
}
```

#### 3.2 字典转换

支持将字典值与字典标签之间进行自动转换，适用于状态、类型等枚举字段。

**第一步：实现字典服务接口**

```java
@Service
public class DictServiceImpl implements DictService {

    @Autowired
    private DictMapper dictMapper;

    @Override
    public String getLabel(String dictType, String dictValue) {
        // 从数据库或缓存中查询字典标签
        // 例如：dictType="sys_user_sex", dictValue="1" -> 返回 "男"
        return dictMapper.selectLabelByTypeAndValue(dictType, dictValue);
    }

    @Override
    public String getValue(String dictType, String dictLabel) {
        // 从数据库或缓存中查询字典值
        // 例如：dictType="sys_user_sex", dictLabel="男" -> 返回 "1"
        return dictMapper.selectValueByTypeAndLabel(dictType, dictLabel);
    }
}
```

**第二步：在实体类中使用**

```java
@Data
public class UserDTO {
    @ExcelProperty(value = "性别", converter = DictConverter.class)
    @ExcelDict(dictType = "sys_user_sex")
    private String sex;  // 数据库存储：1，Excel显示：男

    @ExcelProperty(value = "状态", converter = DictConverter.class)
    @ExcelDict(dictType = "sys_user_status")
    private String status;  // 数据库存储：0，Excel显示：正常

    // 支持多值字典（逗号分隔）
    @ExcelProperty(value = "角色", converter = DictConverter.class)
    @ExcelDict(dictType = "sys_role", separator = ",")
    private String roles;  // 数据库存储：1,2，Excel显示：管理员,普通用户
}
```

**功能说明**：
- 导出时：自动将字典值（如：1）转换为字典标签（如：男）
- 导入时：自动将字典标签（如：男）转换为字典值（如：1）
- 支持多值字典，使用分隔符分隔（默认逗号）

#### 3.3 数据脱敏

支持对敏感数据进行脱敏处理，仅在导出时生效。

**使用示例**：

```java
@Data
public class UserDTO {
    @ExcelProperty(value = "手机号", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.MOBILE_PHONE)
    private String phone;  // 138****1234

    @ExcelProperty(value = "身份证", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.ID_CARD)
    private String idCard;  // 110101********1234

    @ExcelProperty(value = "邮箱", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.EMAIL)
    private String email;  // a***@example.com

    @ExcelProperty(value = "银行卡", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.BANK_CARD)
    private String bankCard;  // 622202******1234

    @ExcelProperty(value = "姓名", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.NAME)
    private String name;  // 张*

    @ExcelProperty(value = "地址", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.ADDRESS)
    private String address;  // 北京市海淀区****

    // 自定义脱敏规则
    @ExcelProperty(value = "自定义", converter = DesensitizeConverter.class)
    @Desensitize(type = DesensitizeType.CUSTOM, prefixKeep = 2, suffixKeep = 3, maskChar = "#")
    private String custom;  // 保留前2位和后3位，中间用#替换
}
```

**支持的脱敏类型**：

| 类型 | 说明 | 示例 |
|------|------|------|
| `MOBILE_PHONE` | 手机号 | 138****1234 |
| `ID_CARD` | 身份证 | 110101********1234 |
| `EMAIL` | 邮箱 | a***@example.com |
| `BANK_CARD` | 银行卡 | 622202******1234 |
| `NAME` | 姓名 | 张*、欧阳** |
| `ADDRESS` | 地址 | 北京市海淀区**** |
| `FIXED_PHONE` | 固定电话 | 010****12 |
| `CAR_LICENSE` | 车牌号 | 京A****1 |
| `CUSTOM` | 自定义 | 根据参数自定义 |

**注意事项**：
- 脱敏仅在导出时生效，导入时不进行脱敏处理
- 可以通过 `enabled` 参数动态控制是否启用脱敏
- 自定义类型可以指定保留位数和脱敏字符

#### 3.4 设置列宽和行高

```java
@Data
public class UserDTO {
    @ExcelProperty("用户ID")
    @ColumnWidth(10)  // 设置列宽
    private Long id;

    @ExcelProperty("用户名")
    @ColumnWidth(20)
    private String username;

    @ExcelProperty("备注")
    @ColumnWidth(50)
    @ContentRowHeight(30)  // 设置行高
    private String remark;
}
```

### 四、配置说明

#### 4.1 全局配置

可以在 `application.yml` 中进行全局配置：

```yaml
allbs:
  excel:
    # Excel 模板文件路径
    template-path: excel/
    # 是否启用国际化
    i18n-enabled: true
```

#### 4.2 配置属性说明

| 属性 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `allbs.excel.template-path` | String | `excel/` | Excel 模板文件路径 |
| `allbs.excel.i18n-enabled` | Boolean | `false` | 是否启用国际化 |

### 五、常见问题

#### 5.1 如何导出大数据量？

Apache Fesod 本身就支持大数据量导出，建议：
- 使用分页查询，避免一次性加载所有数据到内存
- 考虑使用异步导出，避免阻塞请求
- 使用 `@ExportProgress` 注解监听导出进度，提升用户体验

#### 5.2 如何处理日期格式？

使用 `@DateTimeFormat` 注解：

```java
@ExcelProperty("创建时间")
@DateTimeFormat("yyyy-MM-dd HH:mm:ss")
private LocalDateTime createTime;
```

#### 5.3 如何处理数字格式？

使用 `@NumberFormat` 注解：

```java
@ExcelProperty("金额")
@NumberFormat("#.##")
private BigDecimal amount;
```

#### 5.4 Spring Boot 2.x 和 3.x 兼容性

本库同时支持 Spring Boot 2.x 和 3.x，无需任何额外配置。内部已自动处理 `javax.*` 和 `jakarta.*` 包的兼容性。

### 六、嵌套对象导出增强功能 🆕

allbs-excel 提供了三种强大的注解来处理复杂的嵌套对象和列表数据导出。

#### 6.1 功能概览

| 注解 | 适用场景 | 主要功能 |
|------|---------|---------|
| `@NestedProperty` | 需要从嵌套对象中提取单个或多个字段 | 字段路径提取，支持对象、集合、Map |
| `@FlattenProperty` | 需要将整个嵌套对象的所有字段展开 | 自动展开对象的所有 @ExcelProperty |
| `@FlattenList` | 需要将 List 集合展开为多行 | 自动展开 List，支持单元格合并 |

#### 6.2 @NestedProperty - 嵌套对象字段提取

从嵌套对象、集合、Map 中提取指定字段值导出。

**基本用法**：

```java
@Data
public class User {
    @ExcelProperty("用户ID")
    private Long id;

    @ExcelProperty("姓名")
    private String name;

    // 提取部门名称
    @ExcelProperty(value = "部门名称", converter = NestedObjectConverter.class)
    @NestedProperty("name")
    private Department dept;

    // 多层嵌套 - 提取部门领导的姓名
    @ExcelProperty(value = "部门领导", converter = NestedObjectConverter.class)
    @NestedProperty("leader.name")
    private Department dept2;

    // 访问集合第一个元素
    @ExcelProperty(value = "主要技能", converter = NestedObjectConverter.class)
    @NestedProperty("skills[0]")
    private List<String> mainSkill;

    // 拼接所有元素
    @ExcelProperty(value = "所有技能", converter = NestedObjectConverter.class)
    @NestedProperty(value = "skills[*]", separator = ",")
    private List<String> allSkills;

    // 访问 Map 键值
    @ExcelProperty(value = "城市", converter = NestedObjectConverter.class)
    @NestedProperty("properties[city]")
    private Map<String, Object> city;
}
```

**路径表达式语法**：

| 语法 | 说明 | 示例 |
|------|------|------|
| `field` | 访问对象字段 | `dept.name` |
| `field1.field2` | 多层嵌套 | `dept.leader.name` |
| `list[0]` | 访问集合第 N 个元素 | `skills[0]` |
| `list[*]` | 访问集合所有元素并拼接 | `skills[*]` |
| `map[key]` | 访问 Map 指定键的值 | `properties[city]` |

**注解参数**：

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `value` | String | - | 嵌套字段路径表达式（必填） |
| `nullValue` | String | "" | 字段为 null 时的默认值 |
| `separator` | String | "," | 集合元素拼接分隔符 |
| `maxJoinSize` | int | 0 | 集合最大拼接数量，0 表示不限制 |
| `ignoreException` | boolean | true | 是否忽略访问异常 |

#### 6.3 @FlattenProperty - 嵌套对象自动展开

自动展开嵌套对象的所有 `@ExcelProperty` 字段，无需逐个指定路径。

**基本用法**：

```java
@Data
public class User {
    @ExcelProperty("员工ID")
    private Long id;

    @ExcelProperty("员工姓名")
    private String name;

    // 自动展开部门的所有 @ExcelProperty 字段
    @FlattenProperty(prefix = "部门-")
    private Department department;

    // 自动展开上级部门，使用不同的前缀避免冲突
    @FlattenProperty(prefix = "上级部门-")
    private Department parentDept;
}

@Data
public class Department {
    @ExcelProperty("部门编码")
    private String code;

    @ExcelProperty("部门名称")
    private String name;

    @ExcelProperty("部门类型")
    private String type;

    private String internalId;  // 无 @ExcelProperty，不会被导出
}
```

**导出结果**：

| 员工ID | 员工姓名 | 部门-部门编码 | 部门-部门名称 | 部门-部门类型 | 上级部门-部门编码 | 上级部门-部门名称 | 上级部门-部门类型 |
|--------|---------|--------------|--------------|--------------|----------------|----------------|----------------|
| 1 | 张三 | TECH | 技术部 | 研发 | IT | IT中心 | 支持 |

**注解参数**：

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `prefix` | String | "" | 字段名前缀 |
| `suffix` | String | "" | 字段名后缀 |
| `recursive` | boolean | false | 是否递归展开 |
| `maxDepth` | int | 3 | 最大递归深度 |

#### 6.4 @FlattenList - List 实体展开

将 List 集合展开为多行，自动合并单元格。

**基本用法**：

```java
@Data
public class Order {
    @ExcelProperty("订单号")
    private String orderNo;

    @ExcelProperty("下单时间")
    @DateTimeFormat("yyyy-MM-dd HH:mm:ss")
    private LocalDateTime orderTime;

    // 使用 @FlattenProperty 自动展开客户信息
    @FlattenProperty(prefix = "客户-")
    private Customer customer;

    // 使用 @FlattenList 自动展开订单明细
    @FlattenList(prefix = "商品-")
    private List<OrderItem> items;
}

@Data
public class Customer {
    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("手机号")
    private String phone;
}

@Data
public class OrderItem {
    @ExcelProperty("商品名称")
    private String productName;

    @ExcelProperty("数量")
    private Integer quantity;

    @ExcelProperty("单价")
    private BigDecimal price;
}
```

**导出代码**：

```java
@GetMapping("/export-order")
public void exportOrder(HttpServletResponse response) throws IOException {
    // 1. 获取原始数据
    List<Order> orders = orderService.findAll();

    // 2. 展开 List
    List<Map<String, Object>> expandedData = ListEntityExpander.expandData(orders);

    // 3. 生成元数据
    ListEntityExpander.ListExpandMetadata metadata =
        ListEntityExpander.analyzeClass(Order.class);

    // 4. 生成合并区域
    List<ListEntityExpander.MergeRegion> mergeRegions =
        ListEntityExpander.generateMergeRegions(expandedData, metadata);

    // 5. 生成表头
    List<String> headers = ListEntityExpander.generateHeaders(metadata);
    List<List<String>> head = headers.stream()
        .map(Collections::singletonList)
        .collect(Collectors.toList());

    // 6. 设置响应
    response.setContentType("application/vnd.ms-excel");
    response.setCharacterEncoding("utf-8");
    String fileName = URLEncoder.encode("订单列表", "UTF-8");
    response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

    // 7. 导出
    FesodSheet.write(response.getOutputStream())
        .head(head)
        .registerWriteHandler(new ListMergeCellWriteHandler(mergeRegions))
        .sheet("订单列表")
        .doWrite(expandedData);
}
```

**导出结果**：

| 订单号 | 下单时间 | 客户-姓名 | 客户-手机号 | 商品-商品名称 | 商品-数量 | 商品-单价 |
|--------|---------|----------|------------|-------------|----------|----------|
| ORDER001 | 2025-01-01 10:00:00 | 张三 | 138****1234 | iPhone15 | 1 | 5999 |
| ↑（合并） | ↑（合并） | ↑（合并） | ↑（合并） | AirPods Pro | 2 | 1999 |

**多 List 展开策略**：

当一个实体有多个 List 字段时，支持三种策略：

```java
@Data
public class Student {
    @ExcelProperty("学生姓名")
    private String name;

    // MAX_LENGTH: 按最长 List 的长度展开（默认）
    @FlattenList(prefix = "课程-", multiListStrategy = FlattenList.MultiListStrategy.MAX_LENGTH)
    private List<Course> courses;

    @FlattenList(prefix = "奖项-", multiListStrategy = FlattenList.MultiListStrategy.MAX_LENGTH)
    private List<Award> awards;
}
```

| 策略 | 说明 | 适用场景 |
|------|------|---------|
| `MAX_LENGTH` | 按最长 List 的长度展开，短的补空 | 默认推荐 |
| `MIN_LENGTH` | 按最短 List 的长度展开 | 只显示完整数据 |
| `CARTESIAN` | 笛卡尔积展开（慎用） | 需要所有组合 |

**注解参数**：

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `prefix` | String | "" | 字段名前缀 |
| `suffix` | String | "" | 字段名后缀 |
| `multiListStrategy` | Enum | MAX_LENGTH | 多 List 合并策略 |
| `maxRows` | int | 0 | 最大展开行数，0 表示不限制 |
| `mergeCell` | boolean | true | 是否合并单元格 |

**注意事项**：
- `@FlattenList` 需要手动处理导出流程，不能使用 `@ExportExcel` 注解
- List 展开会增加数据量，建议使用 `maxRows` 限制
- 笛卡尔积策略会导致数据量指数增长，慎用

**完整示例参考**：详见 `USAGE.md` 文档。

### 七、样式和表头增强功能 🆕

#### 7.1 @ConditionalStyle - 条件样式

根据单元格值自动应用不同的样式（背景色、字体颜色、加粗等）。

**示例**：

```java
@Data
public class ConditionalStyleDTO {
    @ExcelProperty("学生姓名")
    private String studentName;

    // 根据分数应用不同背景色
    @ExcelProperty("考试分数")
    @ConditionalStyle(conditions = {
        @Condition(value = ">=90", style = @CellStyleDef(backgroundColor = "#00FF00", bold = true)), // 绿色
        @Condition(value = ">=60", style = @CellStyleDef(backgroundColor = "#FFFF00")),             // 黄色
        @Condition(value = "<60", style = @CellStyleDef(backgroundColor = "#FF0000", fontColor = "#FFFFFF")) // 红色白字
    })
    private Integer score;

    // 根据状态应用样式
    @ExcelProperty("任务状态")
    @ConditionalStyle(conditions = {
        @Condition(value = "已完成", style = @CellStyleDef(backgroundColor = "#00FF00", fontColor = "#FFFFFF")),
        @Condition(value = "进行中", style = @CellStyleDef(backgroundColor = "#FFFF00")),
        @Condition(value = "已取消", style = @CellStyleDef(backgroundColor = "#808080", fontColor = "#FFFFFF"))
    })
    private String status;

    // 使用正则表达式匹配
    @ExcelProperty("等级")
    @ConditionalStyle(conditions = {
        @Condition(value = "regex:^A.*", style = @CellStyleDef(backgroundColor = "#00FF00", bold = true)),
        @Condition(value = "regex:^B.*", style = @CellStyleDef(backgroundColor = "#FFFF00")),
        @Condition(value = "regex:^C.*", style = @CellStyleDef(backgroundColor = "#FFA500"))
    })
    private String grade;
}
```

**导出代码**：

```java
// 需要手动注册 ConditionalStyleWriteHandler
FesodSheet.write(response.getOutputStream(), ConditionalStyleDTO.class)
    .registerWriteHandler(new ConditionalStyleWriteHandler(ConditionalStyleDTO.class))
    .sheet("条件样式示例")
    .doWrite(data);
```

**条件表达式支持**：

| 表达式类型 | 格式 | 示例 |
|----------|------|------|
| 精确匹配 | 直接写值 | `"已完成"` |
| 大于 | `>值` | `">100"` |
| 大于等于 | `>=值` | `">=90"` |
| 小于 | `<值` | `"<60"` |
| 小于等于 | `<=值` | `"<=50"` |
| 区间 | `[min,max]` 或 `(min,max)` | `"[60,90]"` |
| 正则表达式 | `regex:表达式` | `"regex:^A.*"` |

**参数说明**：

`@ConditionalStyle` 参数：

| 参数 | 类型 | 默认值 | 说明 |
|-----|------|--------|------|
| `conditions` | Condition[] | 必填 | 条件列表 |
| `enabled` | boolean | true | 是否启用 |

`@Condition` 参数：

| 参数 | 类型 | 默认值 | 说明 |
|-----|------|--------|------|
| `value` | String | 必填 | 条件表达式 |
| `style` | CellStyleDef | 必填 | 应用的样式 |
| `priority` | int | 0 | 优先级（越小越高） |

`@CellStyleDef` 参数：

| 参数 | 类型 | 默认值 | 说明 |
|-----|------|--------|------|
| `foregroundColor` | String | "" | 前景色（#RRGGBB） |
| `backgroundColor` | String | "" | 背景色（#RRGGBB） |
| `fontColor` | String | "" | 字体颜色（#RRGGBB） |
| `bold` | boolean | false | 是否加粗 |
| `fontSize` | short | -1 | 字体大小 |
| `horizontalAlignment` | short | -1 | 水平对齐（1=LEFT, 2=CENTER, 3=RIGHT） |
| `verticalAlignment` | short | -1 | 垂直对齐（0=TOP, 1=CENTER, 2=BOTTOM） |

#### 7.2 @DynamicHeaders - 动态表头

根据数据动态生成表头列，适用于属性不固定的场景（如自定义字段、EAV模型）。

**示例**：

```java
@Data
public class DynamicHeaderDTO {
    @ExcelProperty("产品ID")
    private Long productId;

    @ExcelProperty("产品名称")
    private String productName;

    // 动态表头：从数据中自动提取
    @DynamicHeaders(
        strategy = DynamicHeaderStrategy.FROM_DATA,
        headerPrefix = "属性-",
        order = DynamicHeaders.SortOrder.ASC
    )
    private Map<String, Object> properties;

    // 预定义表头
    @DynamicHeaders(
        strategy = DynamicHeaderStrategy.FROM_CONFIG,
        headers = {"备注1", "备注2", "备注3"},
        headerPrefix = "扩展-"
    )
    private Map<String, Object> extFields;
}
```

**导出代码**：

```java
// 1. 获取数据
List<DynamicHeaderDTO> products = getProducts();

// 2. 展开动态字段
DynamicHeaderProcessor.DynamicHeaderMetadata metadata =
    DynamicHeaderProcessor.analyzeClass(DynamicHeaderDTO.class, products);
List<Map<String, Object>> expandedData = DynamicHeaderProcessor.expandData(products);

// 3. 生成表头
List<String> headers = DynamicHeaderProcessor.generateHeaders(metadata);
List<List<String>> head = headers.stream()
    .map(Collections::singletonList)
    .collect(Collectors.toList());

// 4. 导出
FesodSheet.write(response.getOutputStream())
    .head(head)
    .sheet("产品列表")
    .doWrite(expandedData);
```

**生成策略**：

| 策略 | 说明 | 使用场景 |
|-----|------|----------|
| `FROM_DATA` | 从数据中自动提取所有键作为表头 | 属性完全动态，无法预知 |
| `FROM_CONFIG` | 使用预定义的表头列表 | 属性固定且已知 |
| `MIXED` | 先使用配置的表头，再补充数据中的其他键 | 有必选字段+可选动态字段 |

**参数说明**：

| 参数 | 类型 | 默认值 | 说明 |
|-----|------|--------|------|
| `strategy` | Enum | FROM_DATA | 表头生成策略 |
| `headers` | String[] | {} | 预定义表头（FROM_CONFIG/MIXED时使用） |
| `headerPrefix` | String | "" | 表头前缀 |
| `headerSuffix` | String | "" | 表头后缀 |
| `order` | Enum | NONE | 排序方式（NONE/ASC/DESC） |
| `maxColumns` | int | -1 | 最大列数限制，-1表示不限制 |
| `enabled` | boolean | true | 是否启用 |

**注意事项**：
- 动态表头需要手动处理导出流程，不能使用 `@ExportExcel` 注解
- 建议使用 `maxColumns` 限制列数，防止数据过多导致性能问题
- 不同数据行的动态字段可以不同，最终表头是所有行的并集

### 八、导入增强功能 🆕

#### 8.1 @NestedProperty 嵌套对象导入

使用 `NestedObjectReadConverter` 可以在导入时自动填充嵌套对象字段。

**示例**：

```java
@Data
public class EmployeeImportDTO {
    @ExcelProperty("员工ID")
    private Long id;

    @ExcelProperty("员工姓名")
    private String name;

    // 导入时自动创建 Department 对象并设置 name 字段
    @ExcelProperty(value = "部门名称", converter = NestedObjectReadConverter.class)
    @NestedProperty("name")
    private Department department;

    // 支持多层嵌套
    @ExcelProperty(value = "直属领导", converter = NestedObjectReadConverter.class)
    @NestedProperty("leader.name")
    private Department department2;
}
```

**导入代码**：

```java
List<EmployeeImportDTO> data = FesodSheet.read(file.getInputStream(),
    EmployeeImportDTO.class, null)
    .sheet()
    .doReadSync();
```

**说明**：
- `NestedObjectReadConverter` 会自动创建嵌套对象实例
- 支持多层嵌套路径（如 `leader.name`）
- 自动进行类型转换（String、Integer、Long、Double、Boolean等）

#### 8.2 @FlattenList 多行聚合导入

使用 `FlattenListReadListener` 可以将多行 Excel 数据聚合回包含 List 的对象。

**示例**：

```java
@Data
public class OrderImportDTO {
    @ExcelProperty("订单号")
    private String orderNo;

    @ExcelProperty("下单时间")
    @DateTimeFormat("yyyy-MM-dd HH:mm:ss")
    private LocalDateTime orderTime;

    @FlattenProperty(prefix = "客户-")
    private Customer customer;

    @FlattenList(prefix = "商品-")
    private List<OrderItem> items;
}
```

**导入代码**：

```java
// 创建聚合监听器
FlattenListReadListener<OrderImportDTO> listener =
    new FlattenListReadListener<>(OrderImportDTO.class);

// 读取 Excel
FesodSheet.read(file.getInputStream(), listener)
    .sheet()
    .doRead();

// 获取聚合后的结果
List<OrderImportDTO> result = listener.getResult();
```

**工作原理**：
1. 监听器读取每一行数据
2. 通过普通字段（非 List 字段）判断是否属于同一个对象
3. 如果是同一个对象，将当前行的 List 元素添加到该对象的 List 中
4. 如果是新对象，保存上一个对象并创建新对象
5. 最后返回聚合后的对象列表

**注意事项**：
- Excel 中同一个对象的多行数据必须连续
- 普通字段（如订单号、下单时间等）在同一对象的多行中必须相同
- List 字段的表头需要使用前缀（如 "商品-名称"、"商品-数量"）

#### 8.3 导入导出完整示例

```java
// 1. 导出
@GetMapping("/export-order")
public void exportOrder(HttpServletResponse response) throws IOException {
    List<FlattenListOrderDTO> orders = orderService.getOrders();

    // 展开 List
    List<Map<String, Object>> expandedData = ListEntityExpander.expandData(orders);
    ListEntityExpander.ListExpandMetadata metadata =
        ListEntityExpander.analyzeClass(FlattenListOrderDTO.class);

    // 生成表头
    List<String> headers = ListEntityExpander.generateHeaders(metadata);
    List<List<String>> head = headers.stream()
        .map(Collections::singletonList)
        .collect(Collectors.toList());

    // 设置响应
    response.setContentType("application/vnd.ms-excel");
    response.setCharacterEncoding("utf-8");
    String fileName = URLEncoder.encode("订单明细", "UTF-8");
    response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

    // 导出
    FesodSheet.write(response.getOutputStream())
        .head(head)
        .sheet("订单明细")
        .doWrite(expandedData);
}

// 2. 导入
@PostMapping("/import-order")
public Map<String, Object> importOrder(@RequestParam("file") MultipartFile file) throws IOException {
    // 使用聚合监听器
    FlattenListReadListener<FlattenListOrderDTO> listener =
        new FlattenListReadListener<>(FlattenListOrderDTO.class);

    FesodSheet.read(file.getInputStream(), listener)
        .sheet()
        .doRead();

    List<FlattenListOrderDTO> orders = listener.getResult();

    // 保存到数据库
    orderService.saveOrders(orders);

    Map<String, Object> result = new HashMap<>();
    result.put("success", true);
    result.put("count", orders.size());
    result.put("message", "成功导入 " + orders.size() + " 个订单");

    return result;
}
```

### 九、数据验证功能 🆕

#### 9.1 @ExcelValidation - Excel 数据验证

为 Excel 列添加数据验证规则，限制用户输入，确保数据质量。

**支持的验证类型**:
- 下拉列表（LIST）
- 数值范围（NUMBER_RANGE、INTEGER、DECIMAL）
- 日期验证（DATE、TIME）
- 文本长度（TEXT_LENGTH）
- 自定义公式（FORMULA）
- 任意值（ANY，仅用于提示）

**基本使用**:

```java
@Data
public class EmployeeDTO {
    @ExcelProperty("姓名")
    @ExcelValidation(
        type = ValidationType.TEXT_LENGTH,
        minLength = 2,
        maxLength = 10,
        errorMessage = "姓名长度必须在2-10个字符之间",
        promptMessage = "请输入2-10个字符的姓名",
        showPromptBox = true
    )
    private String name;

    @ExcelProperty("性别")
    @ExcelValidation(
        type = ValidationType.LIST,
        options = {"男", "女"},
        errorMessage = "性别只能选择：男、女"
    )
    private String gender;

    @ExcelProperty("年龄")
    @ExcelValidation(
        type = ValidationType.INTEGER,
        min = 18,
        max = 65,
        errorMessage = "年龄必须在18-65之间"
    )
    private Integer age;

    @ExcelProperty("工资")
    @ExcelValidation(
        type = ValidationType.DECIMAL,
        min = 3000.0,
        max = 50000.0,
        errorMessage = "工资必须在3000-50000之间"
    )
    private Double salary;

    @ExcelProperty("入职日期")
    @ExcelValidation(
        type = ValidationType.DATE,
        dateFormat = "yyyy-MM-dd",
        errorMessage = "请输入有效的日期格式"
    )
    private LocalDate hireDate;
}
```

**导出时应用验证规则**:

```java
@GetMapping("/export/validation")
public void exportWithValidation(HttpServletResponse response) throws IOException {
    List<EmployeeDTO> data = employeeService.findAll();

    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    response.setCharacterEncoding("utf-8");
    String fileName = URLEncoder.encode("员工信息", "UTF-8");
    response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

    // 注册数据验证处理器
    FesodSheet.write(response.getOutputStream(), EmployeeDTO.class)
        .sheet("员工信息")
        .registerWriteHandler(new ExcelValidationWriteHandler(EmployeeDTO.class))
        .doWrite(data);
}
```

**自定义验证范围**:

```java
// 验证范围从第 2 行到第 1000 行
new ExcelValidationWriteHandler(EmployeeDTO.class, 1, 1000)
```

**注解属性说明**:

| 属性 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| type | ValidationType | 验证类型 | - |
| options | String[] | 下拉列表选项（LIST 类型使用） | [] |
| min | double | 最小值（数值类型使用） | Double.MIN_VALUE |
| max | double | 最大值（数值类型使用） | Double.MAX_VALUE |
| minLength | int | 最小长度（TEXT_LENGTH 使用） | 0 |
| maxLength | int | 最大长度（TEXT_LENGTH 使用） | Integer.MAX_VALUE |
| dateFormat | String | 日期格式（DATE/TIME 使用） | "yyyy-MM-dd" |
| formula | String | 自定义公式（FORMULA 使用） | "" |
| errorMessage | String | 错误提示消息 | "输入的数据无效" |
| errorTitle | String | 错误提示标题 | "数据验证错误" |
| promptMessage | String | 输入提示消息 | "" |
| promptTitle | String | 输入提示标题 | "输入提示" |
| showErrorBox | boolean | 是否显示错误警告 | true |
| showPromptBox | boolean | 是否显示输入提示 | false |
| enabled | boolean | 是否启用 | true |

### 十、多 Sheet 关联导出 🆕

#### 10.1 @RelatedSheet - 关联 Sheet 导出

将主表和关联明细数据自动导出到不同的 Sheet，并建立关联关系。

**使用场景**:
- 订单与订单明细
- 部门与员工
- 客户与联系人
- 产品与规格

**基本使用**:

```java
// 订单主表
@Data
public class OrderDTO {
    @ExcelProperty("订单号")
    private String orderNo;

    @ExcelProperty("客户名称")
    private String customerName;

    @ExcelProperty("订单金额")
    private BigDecimal totalAmount;

    @ExcelProperty("订单状态")
    private String status;

    @ExcelProperty("创建时间")
    private LocalDateTime createTime;

    @ExcelProperty("明细数量")
    private Integer itemCount;

    // 关联的订单明细（导出到单独的 Sheet）
    @RelatedSheet(
        sheetName = "订单明细",
        relationKey = "orderNo",
        dataType = OrderItemDTO.class,
        createHyperlink = true,
        hyperlinkText = "查看明细"
    )
    private List<OrderItemDTO> items;
}

// 订单明细
@Data
public class OrderItemDTO {
    @ExcelProperty("订单号")
    private String orderNo;

    @ExcelProperty("序号")
    private Integer itemNo;

    @ExcelProperty("商品名称")
    private String productName;

    @ExcelProperty("数量")
    private Integer quantity;

    @ExcelProperty("单价")
    private BigDecimal price;

    @ExcelProperty("小计")
    private BigDecimal subtotal;
}
```

**导出多 Sheet**:

```java
@GetMapping("/export/multi-sheet")
public void exportMultiSheet(HttpServletResponse response) throws IOException {
    List<OrderDTO> orders = orderService.findAll();

    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    response.setCharacterEncoding("utf-8");
    String fileName = URLEncoder.encode("订单及明细", "UTF-8");
    response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

    // 使用 MultiSheetRelationProcessor 导出
    ExcelWriter excelWriter = FesodSheet.write(response.getOutputStream()).build();
    try {
        MultiSheetRelationProcessor.exportWithRelations(
            excelWriter,
            orders,
            "订单",
            OrderDTO.class
        );
    } finally {
        if (excelWriter != null) {
            excelWriter.finish();
        }
    }
}
```

**注解属性说明**:

| 属性 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| sheetName | String | 关联 Sheet 名称 | - |
| relationKey | String | 主表关联键字段名 | - |
| childRelationKey | String | 子表关联字段名（如果与主表不同） | "" |
| createHyperlink | boolean | 是否创建超链接 | true |
| hyperlinkText | String | 超链接显示文本 | "" |
| dataType | Class<?> | 子表数据类型 | Object.class |
| orderBy | String | 子表排序字段 | "" |
| enabled | boolean | 是否启用 | true |

**功能特点**:
- ✅ 自动提取关联数据到独立 Sheet
- ✅ 支持创建超链接跳转
- ✅ 支持一对多关系
- ✅ 支持自定义关联键
- ✅ 灵活的 Sheet 配置

### 十一、更新日志

#### [3.2.0] - 2026-08-14

**升级**:
- ⬆️ 从 FastExcel 1.3.0 迁移到 Apache Fesod 2.0.2-incubating
- ⬆️ Lombok 与 Maven 构建插件升级到当前兼容版本

#### [2.2.0] - 2025-11-17

**新增**:
- ✨ 新增 `@ExcelValidation` 注解 - Excel 数据验证
  - 支持下拉列表、数值范围、整数、小数、日期、时间、文本长度、自定义公式等验证类型
  - 支持自定义错误提示和输入提示
  - 支持自定义验证范围（起始行和结束行）
- ✨ 新增 `@RelatedSheet` 注解 - 多 Sheet 关联导出
  - 将主表和关联数据自动导出到不同 Sheet
  - 支持创建超链接跳转到关联 Sheet
  - 支持一对多关系
  - 支持自定义关联键和子表关联键
- ✨ 新增 `@SheetRelation` 注解 - Sheet 关系配置
  - 支持配置多个 Sheet 之间的关系
  - 支持自动创建目录 Sheet
- ✨ 新增 `ExcelValidationWriteHandler` - 数据验证处理器
  - 自动分析字段上的 `@ExcelValidation` 注解
  - 根据验证类型创建相应的验证约束
  - 应用验证规则到 Excel 单元格
- ✨ 新增 `MultiSheetRelationProcessor` - 多 Sheet 关联处理器
  - 自动提取关联数据到独立 Sheet
  - 创建 Sheet 间的超链接
  - 支持创建目录 Sheet
- ✨ 新增 `ValidationType` 枚举 - 数据验证类型定义
- ✨ 新增 `@NestedProperty` 注解 - 嵌套对象字段提取
  - 支持从嵌套对象、集合、Map、数组中提取字段值
  - 支持多层嵌套对象访问（如：`dept.leader.name`）
  - 支持集合索引访问（如：`skills[0]`）
  - 支持集合全部元素拼接（如：`skills[*]`）
  - 支持 Map 键值访问（如：`properties[city]`）
  - 支持自定义分隔符和最大拼接数量
- ✨ 新增 `@FlattenProperty` 注解 - 嵌套对象自动展开
  - 自动展开嵌套对象的所有 `@ExcelProperty` 字段
  - 支持字段名前缀和后缀
  - 支持递归展开多层嵌套对象
  - 支持最大递归深度控制
- ✨ 新增 `@FlattenList` 注解 - List 实体展开
  - 将 List 集合展开为多行
  - 自动合并重复的单元格
  - 支持多个 List 同时展开
  - 支持三种多 List 合并策略（MAX_LENGTH、MIN_LENGTH、CARTESIAN）
  - 支持最大行数限制
- ✨ 新增 `@ConditionalStyle` 注解 - 条件样式
  - 根据单元格值自动应用不同样式
  - 支持背景色、字体颜色、加粗、对齐方式等样式设置
  - 支持精确匹配、数值比较、区间、正则表达式等条件
  - 支持条件优先级设置
- ✨ 新增 `@DynamicHeaders` 注解 - 动态表头
  - 根据数据动态生成表头列
  - 支持从数据自动提取、预定义、混合三种策略
  - 支持表头前缀、后缀、排序、列数限制
  - 适用于 EAV 模型、自定义字段等场景
- ✨ 新增 `NestedObjectConverter` - 嵌套对象转换器
- ✨ 新增 `FlattenFieldProcessor` - 对象展开字段处理器
- ✨ 新增 `ListEntityExpander` - List 实体展开工具
- ✨ 新增 `NestedFieldResolver` - 嵌套字段解析器
- ✨ 新增 `ConditionalStyleWriteHandler` - 条件样式处理器
- ✨ 新增 `DynamicHeaderProcessor` - 动态表头处理器
- ✨ 新增 `ListMergeCellWriteHandler` - List 合并单元格处理器
- ✨ 新增 `NestedObjectReadConverter` - 嵌套对象导入转换器
  - 支持导入时自动创建嵌套对象
  - 支持多层嵌套路径解析
  - 自动类型转换
- ✨ 新增 `FlattenListReadListener` - List 聚合导入监听器
  - 将多行 Excel 数据聚合回包含 List 的对象
  - 自动识别并分组相关行
  - 支持复杂嵌套结构

**优化**:
- 🔧 将 `ListEntityExpander.analyzeClass()` 方法改为 public，方便外部调用

**文档**:
- 📖 新增 `USAGE.md` 嵌套对象导出完整使用指南
- 📖 更新 `README.md` 添加数据验证、多 Sheet 关联导出、条件样式、动态表头和导入增强功能说明
- 📖 allbs-excel-test 项目新增 11 个测试接口（9 导出 + 2 导入）和完整前端演示页面

#### [3.0.0] - 2025-11-15

**新增**:
- ✨ 支持空数据导出带表头的 Excel
- ✨ 新增 `@Sheet.clazz` 属性用于指定数据类型
- ✨ 同时支持 Spring Boot 2.x 和 3.x
- ✨ 新增字典转换功能（`@ExcelDict` + `DictConverter`）
- ✨ 新增数据脱敏功能（`@Desensitize` + `DesensitizeConverter`）
- ✨ 支持手机号、身份证、邮箱、银行卡等多种脱敏类型
- ✨ 支持自定义脱敏规则
- ✨ 新增 `onlyExcelProperty` 配置，支持只导出有 `@ExcelProperty` 注解的字段
- ✨ 支持非连续的列索引（如：1、2、7、11）
- ✨ 新增合并单元格功能（`@ExcelMerge` + `MergeCellWriteHandler`）
- ✨ 支持同值自动合并，支持依赖关系合并
- ✨ 新增导出进度回调功能（`@ExportProgress` + `ExportProgressListener`）
- ✨ 支持实时监听导出进度，适用于大数据量导出场景
- ✨ 支持与 WebSocket、SSE 等技术结合实现实时进度推送
- ✨ 增强动态文件名功能，新增预定义变量（`#now`, `#today`, `#timestamp`, `#uuid`）
- ✨ 新增自定义函数（`#formatDate()`, `#formatDateTime()`, `#sanitize()`, `#timestamp()`）
- ✨ 支持更丰富的 SpEL 表达式（字符串操作、条件表达式、数学运算、集合操作）

**升级**:
- ⬆️ EasyExcel 升级到 4.0.3
- ⬆️ Lombok 升级到 1.18.36
- ⬆️ 移除 allbs-common 依赖

**修复**:
- 🐛 修复空 List 无法导出的问题
- 🐛 修复 Maven 部署配置问题

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

本项目采用 [Apache License 2.0](LICENSE) 许可证。

## 👨‍💻 作者

- **ChenQi** - [GitHub](https://github.com/chenqi92)

## 🙏 致谢

- [Apache Fesod (Incubating)](https://fesod.apache.org/) - Apache 基金会孵化中的高性能表格处理工具
- [Spring Boot](https://spring.io/projects/spring-boot) - Spring Boot 框架

## 📮 联系方式

- Email: chenqi92104@icloud.com
- GitHub: https://github.com/chenqi92/allbs-excel
