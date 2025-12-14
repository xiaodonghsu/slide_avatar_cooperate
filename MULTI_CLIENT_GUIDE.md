# 多类型WebSocket客户端功能说明

## 功能概述

新增的多类型客户端功能允许系统支持不同类型的WebSocket客户端连接，并根据客户端类型定向发送消息。

## 支持的客户端类型

- **avatar**: 数字人客户端，接收播放控制、文本播报等消息
- **controller**: 控制器客户端，可能用于接收控制状态信息
- **monitor**: 监控客户端，可能用于接收系统状态信息
- **其他自定义类型**: 可以根据需要添加新的客户端类型

## 工作流程

### 1. 客户端连接
当客户端连接到WebSocket服务器时，会立即收到一个类型识别问询消息：

```json
{
    "type": "inquiry",
    "message": "Please identify your client type (e.g., 'avatar', 'controller', 'monitor')",
    "supported_types": ["avatar", "controller", "monitor"]
}
```

### 2. 客户端身份识别
客户端需要回复识别消息：

```json
{
    "type": "identification",
    "client_type": "avatar"
}
```

### 3. 身份确认
服务器收到识别消息后，会发送确认回复：

```json
{
    "type": "identification_confirmed",
    "client_type": "avatar",
    "message": "Client type confirmed as: avatar"
}
```

### 4. 消息定向发送
服务器现在可以指定发送消息给特定类型的客户端：

```python
# 发送给所有已识别的客户端
await handler.send_to_clients(message)

# 发送给特定类型的客户端
await handler.send_to_clients(message, "avatar")

# 发送给所有客户端（包括未识别的）
await handler.send_to_all_clients(message)
```

## API变更

### Handler类新增方法

- `send_to_clients(message, client_type=None)`: 发送消息给指定类型或所有已识别的客户端
- `send_to_all_clients(message)`: 发送消息给所有客户端
- `get_client_type_stats()`: 获取客户端类型统计信息

### Handler类新增属性

- `client_types`: 存储客户端类型信息的字典 `{websocket: client_type}`
- `pending_identification`: 等待识别的客户端集合

## 使用示例

### 客户端实现示例

```python
import asyncio
import websockets
import json

async def avatar_client():
    uri = "ws://localhost:8765"
    async with websockets.connect(uri) as websocket:
        # 等待问询消息
        inquiry = await websocket.recv()
        print(f"收到问询: {inquiry}")
        
        # 发送身份识别
        identification = {
            "type": "identification",
            "client_type": "avatar"
        }
        await websocket.send(json.dumps(identification))
        
        # 等待确认
        confirmation = await websocket.recv()
        print(f"身份确认: {confirmation}")
        
        # 开始正常通信
        while True:
            message = await websocket.recv()
            print(f"收到消息: {message}")
            # 处理消息...

asyncio.run(avatar_client())
```

## 测试

运行测试程序验证功能：

```bash
python test_multi_client.py
```

测试程序会：
1. 创建4个不同类型的测试客户端
2. 连接到服务器并进行身份识别
3. 监听10秒的消息接收
4. 显示每个客户端收到的消息

## 日志输出

服务器会记录详细的客户端管理日志：

```
Client connected: ('127.0.0.1', 54321)
Sent type inquiry to client: ('127.0.0.1', 54321)
Client ('127.0.0.1', 54321) identified as type: avatar
Broadcasting message to 1 clients of type 'avatar': {"tasks": "text", "text": "启动场景scene1", "duration": 2}
```

## 注意事项

1. 未识别类型的客户端不会接收到定向消息
2. 客户端类型区分大小写
3. 服务器会自动清理断开连接的客户端信息
4. 系统支持同时连接多个同类型的客户端

## 兼容性

- 保持了原有的 `send_to_clients(message)` 方法兼容性
- 未指定 `client_type` 参数时，默认发送给所有已识别的客户端
- 现有的数字人客户端代码无需修改，只需在连接时进行类型识别即可