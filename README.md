# allbs-excel

[![License](https://img.shields.io/badge/license-Apache%202-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)
[![Maven Central](https://img.shields.io/maven-central/v/cn.allbs/allbs-excel.svg)](https://search.maven.org/artifact/cn.allbs/allbs-excel)

基于 [EasyExcel](https://github.com/alibaba/easyexcel) 的 Spring Boot Excel 导入导出增强工具，通过注解即可实现 Excel 的导入导出功能。

## ✨ 特性

- 🚀 **简单易用**: 通过注解即可实现 Excel 导入导出
- 📝 **功能丰富**: 支持单/多 Sheet、模板导出、数据验证等
- 🎨 **灵活定制**: 支持自定义转换器、样式处理器
- 🔄 **字典转换**: 支持字典值与标签的自动转换
- 🔐 **数据脱敏**: 支持手机号、身份证等敏感数据脱敏
- 🌍 **国际化支持**: Excel 表头支持国际化
- 🔒 **数据验证**: 导入时自动进行数据校验
- 📊 **空数据导出**: 支持导出只有表头的空 Excel
- ⚡ **高性能**: 基于 EasyExcel 4.0.3，性能优异
- 🔄 **版本兼容**: 同时支持 Spring Boot 2.x 和 3.x

## 📦 依赖要求

- JDK 17+
- Spring Boot 2.7+ 或 3.x

## 🚀 快速开始

### 1. 添加依赖

```xml
<dependency>
    <groupId>cn.allbs</groupId>
    <artifactId>allbs-excel</artifactId>
    <version>3.0.0</version>
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

默认情况下，EasyExcel 会导出所有字段。如果只想导出标注了 `@ExcelProperty` 的字段，可以使用 `onlyExcelProperty` 配置：

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

支持使用 SpEL 表达式动态生成文件名：

```java
@GetMapping("/export-dynamic")
@ExportExcel(
    name = "用户列表-#{#date}",  // 支持 SpEL 表达式
    sheets = @Sheet(sheetName = "用户信息")
)
public List<UserDTO> exportDynamic(@RequestParam String date) {
    return userService.findByDate(date);
}
```

**SpEL 表达式示例**:
- `用户列表-#{#date}.xlsx` - 使用参数中的 date 值
- `报表-#{T(java.time.LocalDate).now()}.xlsx` - 使用当前日期
- `数据-#{#user.name}.xlsx` - 使用对象属性

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
- 当前版本的合并功能基于 EasyExcel 4.0.3 实现

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

EasyExcel 本身就支持大数据量导出，建议：
- 使用分页查询，避免一次性加载所有数据到内存
- 考虑使用异步导出，避免阻塞请求

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

### 六、更新日志

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

- [EasyExcel](https://github.com/alibaba/easyexcel) - 阿里巴巴开源的 Excel 处理工具
- [Spring Boot](https://spring.io/projects/spring-boot) - Spring Boot 框架

## 📮 联系方式

- Email: chenqi92104@icloud.com
- GitHub: https://github.com/chenqi92/allbs-excel
