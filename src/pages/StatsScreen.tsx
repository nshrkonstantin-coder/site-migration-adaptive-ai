import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useNavigate } from 'react-router-dom';
import Icon from '@/components/ui/icon';

const StatsScreen = () => {
  const navigate = useNavigate();

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100 p-4 md:p-8">
      <div className="max-w-4xl mx-auto">
        <Button 
          variant="ghost" 
          onClick={() => navigate('/')}
          className="mb-4"
        >
          <Icon name="ArrowLeft" size={20} className="mr-2" />
          Назад
        </Button>
        
        <Card className="border-none shadow-lg">
          <CardHeader>
            <CardTitle className="text-2xl flex items-center gap-2">
              <Icon name="BarChart3" size={24} />
              Статистика нарушений
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-muted-foreground">
              Статистика и отчёты в разработке...
            </p>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default StatsScreen;
