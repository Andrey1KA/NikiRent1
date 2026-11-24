import "./globals.css";

export const metadata = {
  title: "Коммуналка — NikiRent",
  description: "Квартиры, коммунальные начисления и выплаты собственникам",
};

export default function RootLayout({ children }) {
  return (
    <html lang="ru">
      <body>{children}</body>
    </html>
  );
}
