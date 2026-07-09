#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
FHX Interlock Extractor (CLI)
从DeltaV FHX文件中提取DCC功能块的联锁/允许/强制/AT模拟跟踪条件信息，生成Excel联锁信息表
"""

import os

from core import parse_fhx, generate_excel


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    fhx_path = os.path.join(base_dir, 'fault.fhx')
    output_path = os.path.join(base_dir, '联锁信息表.xlsx')

    if not os.path.exists(fhx_path):
        print(f"错误: 找不到FHX文件 {fhx_path}")
        return

    print(f"[1/3] 读取FHX文件: {fhx_path}")
    print(f"      文件大小: {os.path.getsize(fhx_path) / 1024:.1f} KB")

    print(f"[2/3] 解析DCC联锁配置...")
    instances = parse_fhx(fhx_path)

    total_ilock = sum(inst.interlock.count_real() for inst in instances)
    total_perm = sum(inst.permissive.count_real() for inst in instances)
    total_force = sum(inst.force.count_real() for inst in instances)
    total_at = sum(inst.tracking.count_real() for inst in instances)

    print(f"      含DCC/AT配置的模块实例: {len(instances)} 个")
    print(f"      已配置联锁条件: {total_ilock} 个")
    print(f"      已配置允许条件: {total_perm} 个")
    print(f"      已配置强制条件: {total_force} 个")
    print(f"      已配置模拟跟踪条件: {total_at} 个")

    print(f"[3/3] 生成Excel文件: {output_path}")
    generate_excel(instances, output_path)
    print(f"\n完成! Excel文件已保存至: {output_path}")


if __name__ == '__main__':
    main()
