使用：

---

**API Endpoint**: `/api/get_furniture_list`

**Method**: `GET`

**Description**: 根据用户的uid、cookie和洞天摹数，构造配置文件，查询并生成Excel文件，然后返回该文件。

**Parameters**:

- `uid`: 用户ID，字符串类型，必需。
- `cookie`: 用户cookie，字符串类型，必需。
- `share_code`: 洞天摹数，字符串类型，必需。

**Responses**:

- `200 OK`: 成功响应，返回一个Excel文件。
- `400 Bad Request`: 请求参数缺失或不正确。
- `404 Not Found`: 生成的Excel文件未找到。

**Example**:

请求：

```
GET /api/get_furniture_list?uid=user123&cookie=abc123&share_code=123456
```

响应：

返回一个名为`user123_123456.xlsx`的Excel文件。

---

