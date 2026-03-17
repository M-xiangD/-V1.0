#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电脑电池健康检测程序
支持Windows系统，获取详细的电池信息并计算健康度
"""

import platform
import subprocess
import re
import json
from datetime import datetime
from typing import Dict, List, Optional, Tuple


class BatteryHealthChecker:
    """电池健康检测器类"""

    def __init__(self):
        self.system = platform.system()
        self.battery_info = {}
        self.health_score = 0
        self.assessment = ""
        self.recommendations = []

    def check_system(self) -> bool:
        """检查操作系统是否支持"""
        if self.system != "Windows":
            print(f"错误：当前系统 {self.system} 暂不支持")
            print("本程序仅支持Windows系统")
            return False
        return True

    def get_battery_info_from_powercfg(self) -> Dict:
        """通过powercfg命令获取电池信息"""
        battery_info = {}

        try:
            # 获取电池报告
            print("执行: powercfg /batteryreport")
            result = subprocess.run(
                ["powercfg", "/batteryreport", "/output", "battery-report.html"],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode != 0:
                print(f"powercfg命令执行失败，返回码: {result.returncode}")
                print(f"错误信息: {result.stderr}")
                return {}

            print("电池报告生成成功")

            # 读取生成的HTML文件
            try:
                with open("battery-report.html", "r", encoding="utf-8") as f:
                    html_content = f.read()
            except UnicodeDecodeError:
                try:
                    with open("battery-report.html", "r", encoding="gbk") as f:
                        html_content = f.read()
                except Exception as e:
                    print(f"读取电池报告文件失败: {e}")
                    return {}
            except FileNotFoundError:
                print("电池报告文件未找到")
                return {}
            except Exception as e:
                print(f"读取电池报告文件时发生错误: {e}")
                return {}

            print(f"成功读取电池报告，文件大小: {len(html_content)} 字节")

            # 解析HTML内容提取电池信息
            battery_info = self._parse_battery_report(html_content)

            if not battery_info:
                print("警告：未能从电池报告中提取任何信息")
            else:
                print(f"成功提取 {len(battery_info)} 项电池信息")

        except subprocess.TimeoutExpired:
            print("powercfg命令执行超时")
        except Exception as e:
            print(f"获取电池信息失败: {e}")

        return battery_info

    def _parse_battery_report(self, html_content: str) -> Dict:
        """解析电池报告HTML内容"""
        battery_info = {}

        # 提取设计容量 - 支持中英文
        design_capacity_patterns = [
            r'<span class="label">DESIGN CAPACITY</span></td><td>([\d,]+)\s*mWh',
            r'<span class="label">设计容量</span></td><td>([\d,]+)\s*mWh',
            r'DESIGN CAPACITY[^>]*>([\d,]+)\s*mWh',
            r'设计容量[^>]*>([\d,]+)\s*mWh'
        ]

        for i, pattern in enumerate(design_capacity_patterns):
            match = re.search(pattern, html_content, re.IGNORECASE)
            if match:
                try:
                    capacity_str = match.group(1).replace(',', '')
                    battery_info['design_capacity'] = int(capacity_str)
                    print(f"使用模式 {i+1} 成功提取设计容量: {battery_info['design_capacity']}")
                    break
                except (ValueError, AttributeError):
                    continue

        # 提取完全充电容量 - 支持中英文
        full_charge_patterns = [
            r'<span class="label">FULL CHARGE CAPACITY</span></td><td>([\d,]+)\s*mWh',
            r'<span class="label">完全充电容量</span></td><td>([\d,]+)\s*mWh',
            r'FULL CHARGE CAPACITY[^>]*>([\d,]+)\s*mWh',
            r'完全充电容量[^>]*>([\d,]+)\s*mWh'
        ]

        for i, pattern in enumerate(full_charge_patterns):
            match = re.search(pattern, html_content, re.IGNORECASE)
            if match:
                try:
                    capacity_str = match.group(1).replace(',', '')
                    battery_info['full_charge_capacity'] = int(capacity_str)
                    print(f"使用模式 {i+1} 成功提取完全充电容量: {battery_info['full_charge_capacity']}")
                    break
                except (ValueError, AttributeError):
                    continue

        # 提取循环次数 - 支持中英文
        cycle_count_patterns = [
            r'<span class="label">CYCLE COUNT</span></td><td>\s*(\d+)',
            r'<span class="label">循环计数</span></td><td>\s*(\d+)',
            r'CYCLE COUNT[^>]*>(\d+)',
            r'循环计数[^>]*>(\d+)'
        ]

        for i, pattern in enumerate(cycle_count_patterns):
            match = re.search(pattern, html_content, re.IGNORECASE)
            if match:
                try:
                    cycle_str = match.group(1).replace(',', '')
                    battery_info['cycle_count'] = int(cycle_str)
                    print(f"使用模式 {i+1} 成功提取循环次数: {battery_info['cycle_count']}")
                    break
                except (ValueError, AttributeError):
                    continue

        # 提取电池化学成分
        chemistry_patterns = [
            r'<span class="label">CHEMISTRY</span></td><td>([^<]+)',
            r'<span class="label">化学成分</span></td><td>([^<]+)',
            r'CHEMISTRY[^>]*>([^<]+)',
            r'化学成分[^>]*>([^<]+)'
        ]

        for i, pattern in enumerate(chemistry_patterns):
            match = re.search(pattern, html_content, re.IGNORECASE)
            if match:
                battery_info['chemistry'] = match.group(1).strip()
                print(f"使用模式 {i+1} 成功提取化学成分: {battery_info['chemistry']}")
                break

        # 提取制造商
        manufacturer_patterns = [
            r'<span class="label">MANUFACTURER</span></td><td>([^<]+)',
            r'<span class="label">制造商</span></td><td>([^<]+)',
            r'MANUFACTURER[^>]*>([^<]+)',
            r'制造商[^>]*>([^<]+)'
        ]

        for i, pattern in enumerate(manufacturer_patterns):
            match = re.search(pattern, html_content, re.IGNORECASE)
            if match:
                battery_info['manufacturer'] = match.group(1).strip()
                print(f"使用模式 {i+1} 成功提取制造商: {battery_info['manufacturer']}")
                break

        # 提取序列号
        serial_number_patterns = [
            r'<span class="label">SERIAL NUMBER</span></td><td>([^<]+)',
            r'<span class="label">序列号</span></td><td>([^<]+)',
            r'SERIAL NUMBER[^>]*>([^<]+)',
            r'序列号[^>]*>([^<]+)'
        ]

        for i, pattern in enumerate(serial_number_patterns):
            match = re.search(pattern, html_content, re.IGNORECASE)
            if match:
                battery_info['serial_number'] = match.group(1).strip()
                print(f"使用模式 {i+1} 成功提取序列号: {battery_info['serial_number']}")
                break

        return battery_info

    def get_battery_info_from_wmi(self) -> Dict:
        """通过WMI获取实时电池信息"""
        battery_info = {}

        try:
            import wmi

            c = wmi.WMI()

            # 获取电池基本信息
            battery = c.Win32_Battery()[0]

            battery_info['estimated_charge_remaining'] = battery.EstimatedChargeRemaining
            battery_info['battery_status'] = battery.BatteryStatus
            battery_info['design_capacity'] = battery.DesignCapacity
            battery_info['full_charge_capacity'] = battery.FullChargeCapacity

            if hasattr(battery, 'CycleCount'):
                battery_info['cycle_count'] = battery.CycleCount

            # 获取电池状态描述
            status_map = {
                1: "正在放电",
                2: "已连接电源，正在充电",
                3: "已充满",
                4: "电量不足",
                5: "严重电量不足",
                6: "正在充电",
                7: "正在充电，电量不足",
                8: "未定义",
                9: "未定义",
                10: "已充满",
                11: "部分充电"
            }
            battery_info['status_description'] = status_map.get(battery.BatteryStatus, "未知状态")

        except ImportError:
            print("未安装wmi库，无法获取实时电池信息")
        except Exception as e:
            print(f"通过WMI获取电池信息失败: {e}")

        return battery_info

    def calculate_health_score(self, design_capacity: Optional[int], full_charge_capacity: Optional[int]) -> float:
        """计算电池健康度分数"""
        if design_capacity is None or full_charge_capacity is None:
            return 0.0

        if design_capacity == 0:
            return 0.0

        health_percentage = (full_charge_capacity / design_capacity) * 100
        return round(health_percentage, 2)

    def assess_battery_health(self, health_score: float, cycle_count: Optional[int] = None) -> Tuple[str, List[str]]:
        """评估电池健康状况并提供建议"""
        assessment = ""
        recommendations = []

        # 基于健康度的评估
        if health_score >= 90:
            assessment = "优秀"
            recommendations.append("电池状态非常好，继续保持良好的使用习惯")
        elif health_score >= 80:
            assessment = "良好"
            recommendations.append("电池状态良好，注意避免高温环境和过度放电")
        elif health_score >= 60:
            assessment = "一般"
            recommendations.append("电池容量有所下降，建议检查使用习惯")
            recommendations.append("可以考虑在电量低于20%时充电，避免深度放电")
        elif health_score >= 40:
            assessment = "较差"
            recommendations.append("电池容量明显下降，建议更换电池")
            recommendations.append("注意：电池性能下降可能影响续航时间")
        else:
            assessment = "极差"
            recommendations.append("电池严重老化，强烈建议更换电池")
            recommendations.append("警告：继续使用可能存在安全风险")

        # 基于循环次数的建议
        if cycle_count is not None:
            if cycle_count < 300:
                recommendations.append(f"循环次数 {cycle_count} 次，处于正常范围")
            elif cycle_count < 500:
                recommendations.append(f"循环次数 {cycle_count} 次，电池已使用一段时间")
                recommendations.append("建议关注电池性能变化")
            elif cycle_count < 800:
                recommendations.append(f"循环次数 {cycle_count} 次，电池接近使用寿命")
                recommendations.append("建议准备更换电池")
            else:
                recommendations.append(f"循环次数 {cycle_count} 次，已超过典型使用寿命")
                recommendations.append("建议尽快更换电池以确保使用安全")

        return assessment, recommendations

    def format_capacity(self, capacity_mwh: Optional[int]) -> str:
        """格式化容量显示"""
        if capacity_mwh is None:
            return "未知"

        wh = capacity_mwh / 1000
        return f"{wh:.2f} Wh ({capacity_mwh:,} mWh)"

    def display_battery_info(self):
        """显示电池信息"""
        print("\n" + "="*60)
        print("                    电池健康检测报告")
        print("="*60)

        if not self.battery_info:
            print("\n无法获取电池信息")
            return

        # 基本信息
        print("\n【基本信息】")
        print(f"检测时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        if 'manufacturer' in self.battery_info:
            print(f"制造商: {self.battery_info['manufacturer']}")

        if 'serial_number' in self.battery_info:
            print(f"序列号: {self.battery_info['serial_number']}")

        if 'chemistry' in self.battery_info:
            print(f"化学成分: {self.battery_info['chemistry']}")

        # 容量信息
        print("\n【容量信息】")
        if 'design_capacity' in self.battery_info:
            print(f"设计容量: {self.format_capacity(self.battery_info['design_capacity'])}")

        if 'full_charge_capacity' in self.battery_info:
            print(f"当前完全充电容量: {self.format_capacity(self.battery_info['full_charge_capacity'])}")

        # 容量损失
        if 'design_capacity' in self.battery_info and 'full_charge_capacity' in self.battery_info:
            design_cap = self.battery_info['design_capacity']
            full_charge_cap = self.battery_info['full_charge_capacity']

            if design_cap is not None and full_charge_cap is not None:
                capacity_loss = design_cap - full_charge_cap
                loss_percentage = (capacity_loss / design_cap) * 100
                print(f"容量损失: {capacity_loss:,} mWh ({loss_percentage:.2f}%)")

        # 循环次数
        if 'cycle_count' in self.battery_info:
            print(f"\n【使用情况】")
            print(f"循环次数: {self.battery_info['cycle_count']} 次")

        # 实时状态
        if 'estimated_charge_remaining' in self.battery_info:
            print(f"\n【实时状态】")
            print(f"当前电量: {self.battery_info['estimated_charge_remaining']}%")

        if 'status_description' in self.battery_info:
            print(f"电池状态: {self.battery_info['status_description']}")

        # 健康评估
        print("\n【健康评估】")
        print(f"健康度: {self.health_score}%")
        print(f"评估结果: {self.assessment}")

        # 建议
        print("\n【使用建议】")
        for i, recommendation in enumerate(self.recommendations, 1):
            print(f"{i}. {recommendation}")

        print("\n" + "="*60)

    def save_report(self, filename: str = "battery_health_report.json"):
        """保存检测报告到JSON文件"""
        report = {
            "检测时间": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "电池信息": self.battery_info,
            "健康度": self.health_score,
            "评估结果": self.assessment,
            "建议": self.recommendations
        }

        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2)
            print(f"\n报告已保存到: {filename}")
        except Exception as e:
            print(f"保存报告失败: {e}")

    def run(self, save_report: bool = False):
        """运行电池健康检测"""
        print("正在检测电池信息...")

        if not self.check_system():
            return

        # 获取电池信息
        print("正在通过powercfg获取电池报告...")
        powercfg_info = self.get_battery_info_from_powercfg()
        print(f"powercfg获取到信息: {len(powercfg_info)} 项")

        print("正在通过WMI获取实时电池信息...")
        wmi_info = self.get_battery_info_from_wmi()
        print(f"WMI获取到信息: {len(wmi_info)} 项")

        # 合并信息 - 优先使用WMI的实时信息
        self.battery_info = powercfg_info.copy()
        # 直接添加WMI中的所有信息，包括实时电量
        for key, value in wmi_info.items():
            if value is not None:
                self.battery_info[key] = value

        if not self.battery_info:
            print("错误：无法获取任何电池信息")
            print("\n可能的原因：")
            print("1. 电脑没有电池（台式机）")
            print("2. 电池未正确安装")
            print("3. 需要管理员权限运行程序")
            print("4. 电池驱动程序问题")
            return

        print(f"\n成功获取电池信息: {len(self.battery_info)} 项")

        # 计算健康度
        if 'design_capacity' in self.battery_info and 'full_charge_capacity' in self.battery_info:
            print(f"设计容量: {self.battery_info['design_capacity']} mWh")
            print(f"完全充电容量: {self.battery_info['full_charge_capacity']} mWh")

            self.health_score = self.calculate_health_score(
                self.battery_info['design_capacity'],
                self.battery_info['full_charge_capacity']
            )

            # 评估健康状况
            cycle_count = self.battery_info.get('cycle_count')
            self.assessment, self.recommendations = self.assess_battery_health(
                self.health_score,
                cycle_count
            )
        else:
            print("\n警告：缺少容量信息，无法计算健康度")
            print(f"已获取的信息: {list(self.battery_info.keys())}")

        # 显示结果
        self.display_battery_info()

        # 保存报告
        if save_report:
            self.save_report()


def main():
    """主函数"""
    print("="*60)
    print("           电脑电池健康检测程序 v1.0")
    print("="*60)

    checker = BatteryHealthChecker()

    # 运行检测并保存报告
    checker.run(save_report=True)

    print("\n检测完成！")


if __name__ == "__main__":
    main()
