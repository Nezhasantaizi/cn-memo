# cn-memo

#### 介绍
独立自研记忆复习软件，参考公开间隔重复算法本土化开发；暂未提供多格式导入，可按需拓展开发，项目持续迭代，适配国产系统与中文使用场景。

**项目声明**：cn-memo 是完全独立自研开源记忆复习工具，仅借鉴间隔重复（艾宾浩斯）学习思维，与 Anki 不存在任何隶属、衍生、合作关系，请勿混淆。

**面向人群**：覆盖全学段、全行业背诵学习人群——中小学、大学、研究生、医师/护士/药师/建造师/财会/公考/教资等各类职业考试，以及所有需要长期记忆、反复复习知识点的学习者。

**核心功能**：多格式题库导入导出（docx/xlsx/pdf/odt）、自定义间隔复习算法、多媒体卡片、批量题库处理、跨平台支持（Windows/Linux/macOS/安卓/iOS/鸿蒙）、多端数据同步、国产系统适配（Deepin深度应用商店）。

**当前状态**：项目基础框架已搭建完成，首要攻坚目标为多格式文档素材导入解析功能。

#### 软件架构
- 后端：Python
- 桌面端：Qt
- 前端：Web 前端技术栈
- 移动端：安卓 / 鸿蒙 / iOS 原生开发
- 打包编译：跨平台打包与 CI/CD

#### 安装教程
1. 克隆仓库：`git clone https://gitee.com/ranxiwen/cn-memo.git`
2. 安装依赖：`pip install -r requirements.txt`
3. 启动程序：`python main.py`
4. 桌面端打包安装包（Windows exe / Linux deb&RPM / macOS）开发中，敬请期待

#### 使用说明
1. 启动程序后，创建记忆牌组
2. 导入题库素材或手动录入知识点卡片
3. 根据间隔复习算法提示，每日完成复习任务
4. 详细使用教程将在功能完善后补充

#### 参与贡献
1. 招募方向：Python后端、Qt桌面开发、前端、软件打包编译、安卓/鸿蒙/iOS移动端开发、学科教研人员、社群运营、测试爱好者
2. 参与原则：所有成员自愿参与，完全自由，不强买强卖、不强制分配任务，纯开源兴趣协作，可自由选择投入时间
3. 协作流程：加入QQ群 → 了解项目 → 认领任务 → 提交代码（Gitee Pull Request或群内协调）

#### 交流与联系方式
- **QQ群（主阵地）**：所有开发沟通、需求研讨、问题交流统一在QQ群进行
- 👉 [点击加入QQ群](https://qun.qq.com/universal-share/share?ac=1&authKey=OISxz36eb5tnn%2FXK3Kw2ibRh%2FOUHxjdqr0bBUDSaLiYysuEzpZJ3ebZb%2Fu1lWU&busi_data=eyJncm91cENvZGUiOiI5NDkzOTIwMDIiLCJ0b2tlbiI6IjNIbXBrNWx1RnNxb21aUXo3dkZnTGFKNEl5U1BVV2l2UjEzb1h2THNmcCtYSlRNS21vUysxUDk2TmtwNHd4S0QiLCJ1aW4iOiIzMTczNDMxOTEzIn0%3D&data=GjEYFOtJP3Uvl5Dlga4TvyYJhdiqx-nmN1W9NapB6fckrzhkv1Hp6kzo6RNHGL4T4V82cW8s6dxE8g4fKR1wcg&svctype=4&tempid=h5_group_info)
- 手机号：19112749580
- 微信号：ffrfstgr6652366gy6
- QQ邮箱：3173431913@qq.com
- Gitee仓库：[https://gitee.com/ranxiwen/cn-memo.git](https://gitee.com/ranxiwen/cn-memo.git)

#### 相关文档
- [免责声明](DISCLAIMER.md)
- [贡献指南](CONTRIBUTING.md)
- [更新日志](CHANGELOG.md)
- [安全策略](SECURITY.md)