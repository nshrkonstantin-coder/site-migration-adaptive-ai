import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { format } from 'date-fns';
import { ru } from 'date-fns/locale';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Avatar, AvatarFallback } from '@/components/ui/avatar';
import Icon from '@/components/ui/icon';

const HomeScreen = () => {
  const navigate = useNavigate();
  const [currentTime, setCurrentTime] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentTime(new Date());
    }, 1000);

    return () => clearInterval(timer);
  }, []);

  const userProfile = {
    fullName: 'Константин Ншрк',
    company: 'ООО "Промбезопасность"',
    position: 'Инженер по технике безопасности',
    email: 'nshrkonstantin@gmail.com'
  };

  const getInitials = (name: string) => {
    return name
      .split(' ')
      .map(n => n[0])
      .join('')
      .toUpperCase()
      .slice(0, 2);
  };

  const mainActions = [
    {
      title: 'Регистрация нарушений',
      icon: 'AlertTriangle',
      color: 'from-red-500 to-red-600',
      route: '/register-violation'
    },
    {
      title: 'Просмотр моих нарушений',
      icon: 'FileText',
      color: 'from-blue-500 to-blue-600',
      route: '/my-violations'
    },
    {
      title: 'Статистика нарушений',
      icon: 'BarChart3',
      color: 'from-green-500 to-green-600',
      route: '/stats'
    }
  ];

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        <Card className="border-none shadow-lg">
          <CardHeader className="flex flex-row items-center gap-4 pb-4">
            <Avatar className="h-16 w-16 border-2 border-primary">
              <AvatarFallback className="text-lg font-semibold bg-primary text-primary-foreground">
                {getInitials(userProfile.fullName)}
              </AvatarFallback>
            </Avatar>
            <div className="flex-1">
              <CardTitle className="text-2xl mb-1">{userProfile.fullName}</CardTitle>
              <p className="text-sm text-muted-foreground">{userProfile.position}</p>
              <p className="text-sm text-muted-foreground">{userProfile.company}</p>
            </div>
          </CardHeader>
          <CardContent>
            <div className="flex items-center gap-2 text-muted-foreground">
              <Icon name="Mail" size={16} />
              <span className="text-sm">{userProfile.email}</span>
            </div>
          </CardContent>
        </Card>

        <Card className="border-none shadow-lg overflow-hidden">
          <div className="bg-gradient-to-r from-primary to-primary/80 text-primary-foreground p-6 text-center">
            <div className="flex items-center justify-center gap-2 mb-2">
              <Icon name="Clock" size={20} />
              <p className="text-sm font-medium opacity-90">Текущее время</p>
            </div>
            <p className="text-4xl font-bold tabular-nums">
              {format(currentTime, 'HH:mm:ss', { locale: ru })}
            </p>
            <p className="text-sm opacity-90 mt-2">
              {format(currentTime, 'EEEE, d MMMM yyyy', { locale: ru })}
            </p>
          </div>
        </Card>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {mainActions.map((action) => (
            <button
              key={action.route}
              onClick={() => navigate(action.route)}
              className="group relative overflow-hidden rounded-2xl p-8 text-white shadow-xl transition-all duration-300 hover:shadow-2xl hover:-translate-y-1 active:translate-y-0 active:shadow-lg"
              style={{
                background: `linear-gradient(135deg, var(--tw-gradient-stops))`,
              }}
            >
              <div className={`absolute inset-0 bg-gradient-to-br ${action.color} opacity-100`} />
              
              <div className="relative z-10 flex flex-col items-center text-center space-y-4">
                <div className="p-4 bg-white/20 rounded-full backdrop-blur-sm transition-transform duration-300 group-hover:scale-110">
                  <Icon name={action.icon as any} size={32} />
                </div>
                <h3 className="text-lg font-semibold leading-tight">
                  {action.title}
                </h3>
              </div>

              <div className="absolute inset-0 bg-gradient-to-t from-black/20 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-300" />
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};

export default HomeScreen;
