/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export default function App() {
  return (
    <main className="w-full h-screen overflow-hidden bg-[#f3f6f2]">
      <iframe
        src="/simulator.html"
        title="토마토 스마트팜 마스터플랜 및 생리 보정 시뮬레이터"
        className="w-full h-full border-0 block"
      />
    </main>
  );
}
