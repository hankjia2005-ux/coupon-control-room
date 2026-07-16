import type { Metadata } from "next";
import initialRows from "../public/initial-data.json";
import DashboardClient, { type CouponRow } from "./components/DashboardClient";

export const metadata: Metadata = {
  title: "东券禁券监查 · 今日看板",
  description: "每日优惠券禁券监查在线共享看板",
};

export default function Home() {
  return (
    <DashboardClient
      initialRows={initialRows as unknown as CouponRow[]}
      initialSource="初始数据 · 2026-07-16"
      initialUpdatedAt="2026-07-16 18:47"
      canEdit={false}
      userName="未登录"
    />
  );
}
